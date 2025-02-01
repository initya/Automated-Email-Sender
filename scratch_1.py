import smtplib
import pandas as pd
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# Load the Excel file
file_path = r"C:\Users\cnity\Downloads\CompanyWise HR contact (1).xlsx"  # Update path if needed
df = pd.read_excel(file_path)

# Email settings
SMTP_SERVER = "smtp.gmail.com"  # Change if using Outlook/Yahoo
SMTP_PORT = 587
EMAIL_ADDRESS = "test.4@gmail.com"  # Your email
EMAIL_PASSWORD = "ofkv bclg khdf jtbl"  # Use an app password if required
RESUME_PATH = r"C:\Users\cnity\Downloads\resume.pdf"  # Update with actual resume path

# Create SMTP session
server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
server.starttls()
server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)

# Iterate over contacts
for index, row in df.iterrows():
    name = row.get("Name", "HR")  # Default fallback
    email = row.get("Email")
    title = row.get("Title", "")
    company = row.get("Company", "")

    if pd.isna(email):  # Skip if email is missing
        continue

    # Compose email
    msg = MIMEMultipart()
    msg["From"] = EMAIL_ADDRESS
    msg["To"] = email
    msg["Subject"] = f"Job Inquiry - {title} at {company}"

    body = f"""Dear {name},

I hope this email finds you well. I am a 3rd year CSE student and I am writing to express my interest in opportunities at {company}. Please find my resume attached for your reference.

Looking forward to hearing from you.

Best regards,  
Name
Ph. No.
any other details
"""
    msg.attach(MIMEText(body, "plain"))

    # Attach resume
    with open(RESUME_PATH, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(RESUME_PATH)}")
        msg.attach(part)

    # Send email
    server.sendmail(EMAIL_ADDRESS, email, msg.as_string())
    print(f"Email sent to {name} ({email})")

# Close SMTP session
server.quit()
print("All emails sent successfully.")
