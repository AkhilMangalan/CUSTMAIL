import os, re, json, base64
import pandas as pd
import requests
from docxtpl import DocxTemplate
from docx2pdf import convert
from PyPDF2 import PdfReader, PdfWriter

def parse_multi(value):
    if not value:
        return []
    return [p.strip() for p in re.split(r"[;,]", str(value)) if p.strip()]

def ensure_list(x):
    return [] if not x else x if isinstance(x, list) else [x]

def set_pdf_metadata(pdf_path, title, author):
    reader = PdfReader(pdf_path)
    writer = PdfWriter()
    for p in reader.pages:
        writer.add_page(p)
    writer.add_metadata({"/Title": title, "/Author": author})
    with open(pdf_path, "wb") as f:
        writer.write(f)

def send_email(api_key, sender_email, sender_name,
               to_email, subject, body, attachment,
               cc=None, bcc=None):

    headers = {
        "accept": "application/json",
        "content-type": "application/json",
        "authorization": api_key
    }

    payload = {
        "from": {"address": sender_email, "name": sender_name},
        "to": [{"email_address": {"address": e}} for e in ensure_list(to_email)],
        "subject": subject,
        "htmlbody": body.replace("\n","<br>")
    }

    if cc: payload["cc"] = [{"email_address":{"address":e}} for e in ensure_list(cc)]
    if bcc: payload["bcc"] = [{"email_address":{"address":e}} for e in ensure_list(bcc)]

    with open(attachment,"rb") as f:
        payload["attachments"] = [{
            "name": os.path.basename(attachment),
            "mime_type":"application/pdf",
            "content": base64.b64encode(f.read()).decode()
        }]

    r = requests.post("https://api.zeptomail.in/v1.1/email",
                      headers=headers, data=json.dumps(payload), timeout=30)
    if r.status_code not in (200,202):
        raise Exception(r.text)

def process_and_send(excel, template, api_key,
                     sender_email, sender_name,
                     subject, body, cc, bcc):

    out = "Output/PDF"
    os.makedirs(out, exist_ok=True)

    df = pd.read_excel(excel)
    logs = []

    for _, row in df.iterrows():
        try:
            context = row.fillna("").to_dict()
            name = str(context.get("Airline_Name","Mail")).replace("/","-")
            pdf = os.path.join(out, f"{name}.pdf")

            doc = DocxTemplate(template)
            doc.render(context)
            doc.save("temp.docx")
            convert("temp.docx", pdf)

            set_pdf_metadata(pdf, subject, sender_name)
            email_body = body.format(**context)

            send_email(api_key, sender_email, sender_name,
                       row["Email"], subject, email_body, pdf,
                       parse_multi(cc), parse_multi(bcc))

            logs.append(f"Sent → {row['Email']}")
        except Exception as e:
            logs.append(f"Failed → {row.get('Email')} : {e}")

    return logs
