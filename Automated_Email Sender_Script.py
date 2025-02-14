import os
import pandas as pd
import ssl
import smtplib
import time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import random

# Function to send an email with optional attachments
def send_email(receiver_email, subject, message, attachments=None):#remove None to send with attachments
    sender_email = "1234@gmail.com"  # Sender's email address
    sender_password = "1234"  # Sender's email password
    smtp_server = "smtp.gmail.com"  # SMTP server for Gmail
    smtp_port = 465  # Port for SSL connection

    # Create a multipart email message
    msg = MIMEMultipart()
    msg["Subject"] = subject
    msg["From"] = sender_email
    msg["To"] = receiver_email
    msg.attach(MIMEText(message, "html"))

    # Attach files if provided
    if attachments:
        for file_path in attachments:
            try:
                with open(file_path, "rb") as attachment_file:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(attachment_file.read())
                encoders.encode_base64(part)
                part.add_header(
                    "Content-Disposition",
                    f"attachment; filename= {os.path.basename(file_path)}",
                )
                msg.attach(part)
            except Exception as e:
                print(f"Could not attach file {file_path}: {e}")

    # Securely connect to the SMTP server and send the email
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(smtp_server, smtp_port, context=context) as server:
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, receiver_email, msg.as_string())

# Function to update log file with sent email records
def update_log(log_file, data):
    with open(log_file, "a") as log:
        log.write(data + "\n")

# Function to read log file and return a list of previously sent emails
def read_log(log_file):
    try:
        with open(log_file, "r") as log:
            data = log.read().splitlines()
        return data
    except FileNotFoundError:
        return []

# Define the path to the Excel file containing recipient information
excel_file_path = r"path/to/.xlxs"
xls = pd.ExcelFile(excel_file_path)
sheet_names = xls.sheet_names  # Get all sheet names in the Excel file

# Log files to track sent emails and errors
log_file = "sent_emails.txt"
error_file = "error_log.csv"
error_data = []
sent_emails_log = read_log(log_file)  # Retrieve previously sent emails

# Iterate through each sheet in the Excel file
for sheet_name in sheet_names:
    df = pd.read_excel(excel_file_path, sheet_name=sheet_name)
    for _, row in df.iterrows():
        # Extract recipient details from each row
        Reciever = row["Reciever"]
        Reciever_email = row["Reciever Email"]
        user_name = row["User name"]
        password = row["Password"]

        # Skip sending if the email was already sent
        if Reciever_email in sent_emails_log:
            print(f"Email already sent to {Reciever}, {Reciever_email}. Skipping...")
            continue

        # Define the email subject and body
        subject = "Subject"
        html = f"""
        <html>
        <head>
            <title>Credentials</title>
            <meta name="viewport" content="width=device-width, initial-scale=1">
        </head>
        <body>
            <div>
                <p><b>Dear {Reciever},</b></p>
                <p>I hope this email finds you well.</p>
                <p>Below are the credentials and access details to get you all set up.</p>
                <p><b>Access Credentials:</b><br>
                   <b>Platform/URL:</b> <a href="https://example.com/">https://example.com/</a><br>
                   <b>Username:</b> {user_name}<br>
                   <b>Password:</b> {password}</p>
                <p>Good luck!</p>
                <p>Best regards,<br>
                   Organization Name</p>
            </div>
        </body>
        </html>
        """

        # Define file attachments (modify paths as needed)
        attachments = [r"path/to/attachment1.pdf", r"path/to/attachment2.jpg"]  # Example file paths

        try:
            # Send the email
            send_email(Reciever_email, subject, html, attachments)
            message = f"{Reciever}, {Reciever_email}"
            print(f"Email sent to {message}")
            update_log(log_file, Reciever_email)  # Log the successful email sending
            time.sleep(random.randint(1, 3))  # Introduce random delay to avoid rate limits
        except Exception as e:
            # Log any errors that occur while sending emails
            error_data.append({
                "Reciever": Reciever,
                "Reciever Email": Reciever_email,
                "Error": str(e)
            })
            error_df = pd.DataFrame(error_data)
            error_df.to_csv(error_file, mode='a', header=not os.path.exists(error_file), index=False)
            print(f"Error sending email to {Reciever}, {Reciever_email}: {e}")
