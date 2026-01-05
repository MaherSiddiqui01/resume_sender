import smtplib
import ssl
import pandas as pd
import time
import random
import csv
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from dotenv import load_dotenv

load_dotenv()
YOUR_NAME = os.getenv("YOUR_NAME")
LINKEDIN_URL = os.getenv("LINKEDIN_URL")
DELAY_SECONDS = 5 * 60
CONTACTS_FILE = "tabahi-ultimate-sheet.xlsx"
LOG_SENT = "sent.csv"
LOG_FAILED = "failed.csv"

SUBJECTS = [
    "Application for Opportunities at Your Organization",
    "Resume for Suitable Roles & Exploring Opportunities",
    "Applying for Relevant Positions - Resume Attached",
    "Interest in Open Roles at Your Organization"
]

BODY_TEMPLATES = [
    """Dear {name},

I hope this email finds you well. My name is {sender}, and I am reaching out to explore any potential opportunities at {org} that align with my background in AI/ML, software development, and problem-solving. I have been following {org}'s work and would be excited about the opportunity to contribute to your team if there is a suitable opening.

I have attached my resume for your consideration. If there are any current or upcoming roles where my skills may be a good fit, I would greatly appreciate the opportunity to connect.

Thank you for your time and consideration.

Best regards,
{sender}
LinkedIn: {linkedin}
""",

    """Dear {name},

I hope you are doing well. I am writing to inquire about any potential opportunities at {org} that may align with my experience in software development and applied AI/ML. I greatly admire the work being done at {org} and would welcome the chance to contribute meaningfully to your team. My resume is attached for your review.

Please feel free to reach out if my profile aligns with your requirements. I would be happy to discuss further.

Kind regards,
{sender}
LinkedIn: {linkedin}
""",

    """Dear {name},

I hope this message finds you well. I am reaching out to express my interest in exploring possible opportunities at {org}. With a strong foundation in software development and problem-solving, I am keen to apply my skills in a professional environment like yours. I have attached my resume for your consideration.

Thank you for your time, and I look forward to the possibility of connecting.

Sincerely,
{sender}
LinkedIn: {linkedin}
"""
]

def load_sent_set():
    if not os.path.exists(LOG_SENT):
        return set()
    s = set()
    with open(LOG_SENT, newline='', encoding='utf-8') as f:
        for row in csv.reader(f):
            if row:
                s.add(row[0].strip().lower())
    return s

def log_sent(email):
    with open(LOG_SENT, 'a', newline='', encoding='utf-8') as f:
        csv.writer(f).writerow([email])

def log_failed(email, reason):
    with open(LOG_FAILED, 'a', newline='', encoding='utf-8') as f:
        csv.writer(f).writerow([email, reason])

def create_email(to_email, subject, body, resume_path=os.getenv("RESUME_FILE")):
    msg = MIMEMultipart()
    msg["From"] = os.getenv("YOUR_EMAIL")
    msg["To"] = to_email
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))
    with open(resume_path, "rb") as f:
        part = MIMEApplication(f.read(), Name=os.getenv("ATTACH_NAME"))
        part['Content-Disposition'] = f'attachment; filename="{os.getenv("ATTACH_NAME")}"'
        msg.attach(part)
    return msg

def send_bulk():
    df = pd.read_excel(CONTACTS_FILE, engine='openpyxl')
    sent_set = load_sent_set()
    smtp_server = "smtp.gmail.com"
    smtp_port = 465
    context = ssl.create_default_context()

    for _, row in df.iterrows():
        org = str(row["Organization Name"]).strip()
        name = str(row["Contact Person"]).strip()
        email = str(row["Email ID"]).strip().lower()

        if not email:
            continue

        if email in sent_set:
            continue

        subject = random.choice(SUBJECTS)
        body = random.choice(BODY_TEMPLATES).format(name=name, org=org, sender=YOUR_NAME, linkedin=LINKEDIN_URL)
        msg = create_email(email, subject, body)

        print(f"\nSending to: {email} — {name} @ {org}")

        try:
            with smtplib.SMTP_SSL(smtp_server, smtp_port, context=context) as server:
                server.login(os.getenv("YOUR_EMAIL"), os.getenv("YOUR_APP_PASSWORD"))
                server.sendmail(os.getenv("YOUR_EMAIL"), email, msg.as_string())

            print(f"Sent to {email}")
            log_sent(email)
            print(f"Waiting {DELAY_SECONDS//60} minutes...\n")
            time.sleep(DELAY_SECONDS)

        except Exception as e:
            print(f"Failed to send to {email}: {e}")
            log_failed(email, str(e))
            print("Skipping wait due to failure. Moving to next.\n")
            continue

if __name__ == "__main__":
    send_bulk()
