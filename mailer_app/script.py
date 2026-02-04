import os
import re
import json
import base64
import pandas as pd
import requests
from docxtpl import DocxTemplate
from docx2pdf import convert
from PyPDF2 import PdfReader, PdfWriter

def parse_multi(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return []
    return [p.strip() for p in re.split(r"[;,]", str(value)) if p.strip()]

def ensure_list(x):
    return [] if not x else x if isinstance(x, list) else [x]

def set_pdf_metadata(pdf_path, title, author):
    reader = PdfReader(pdf_path)
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)
    writer.add_metadata({"/Title": title, "/Author": author})
    with open(pdf_path, "wb") as f:
        writer.write(f)

def send_email_zeptomail(api_key, sender_email, sender_name,
                         to_email, subject, body, attachment_path,
                         cc_emails=None, bcc_emails=None):

    headers = {
        "accept": "application/json",
        "content-type": "application/json",
        "authorization": api_key
    }

    payload = {
        "from": {"address": sender_email, "name": sender_name},
        "to": [{"email_address": {"address": e}} for e in ensure_list(to_email)],
        "subject": subject,
        "htmlbody": body.replace("\n", "<br>")
    }

    if cc_emails:
        payload["cc"] = [{"email_address": {"address": e}} for e in ensure_list(cc_emails)]
    if bcc_emails:
        payload["bcc"] = [{"email_address": {"address": e}} for e in ensure_list(bcc_emails)]

    if attachment_path:
        with open(attachment_path, "rb") as f:
            payload["attachments"] = [{
                "name": os.path.basename(attachment_path),
                "mime_type": "application/pdf",
                "content": base64.b64encode(f.read()).decode()
            }]

    r = requests.post("https://api.zeptomail.in/v1.1/email",
                      headers=headers, data=json.dumps(payload), timeout=30)

    if r.status_code not in (200, 202):
        raise Exception(r.text)

def process_and_send(excel_path, template_path, api_key,
                     sender_email, sender_name,
                     subject, body,
                     cc_default, bcc_default):

    base_dir = os.path.dirname(os.path.abspath(__file__))
    pdf_dir = os.path.join(base_dir, "Output", "PDF")
    os.makedirs(pdf_dir, exist_ok=True)

    df = pd.read_excel(excel_path)
    results = []

    for _, row in df.iterrows():
        try:
            context = row.fillna("").to_dict()

            file_name = str(context.get("Airline_Name","Mail")).replace("/","-")
            pdf_path = os.path.join(pdf_dir, f"{file_name}.pdf")

            doc = DocxTemplate(template_path)
            doc.render(context)
            doc.save("temp.docx")
            convert("temp.docx", pdf_path)

            set_pdf_metadata(pdf_path, subject, sender_name)

            email_body = body.format(**context)

            send_email_zeptomail(
                api_key, sender_email, sender_name,
                row["Email"], subject, email_body, pdf_path,
                parse_multi(cc_default), parse_multi(bcc_default)
            )

            results.append(f"Sent to {row['Email']}")

        except Exception as e:
            results.append(f"Failed {row.get('Email')} : {e}")

    return results
