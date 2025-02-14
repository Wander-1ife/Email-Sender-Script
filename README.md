# Email Sender Script

## Overview
This Python script automates the process of sending emails with attachments to multiple recipients using data from an Excel file. It ensures that emails are not sent multiple times to the same recipients by maintaining a log file. Additionally, it logs any errors that occur during the email-sending process.

## Features
- Reads recipient details from an Excel file
- Sends emails with custom messages and attachments
- Logs sent emails to prevent duplicate sending
- Handles errors and logs them for debugging
- Introduces random delays to avoid email server throttling

## Prerequisites
- Python 3.x
- Required libraries:
  - `pandas`
  - `smtplib`
  - `email`
  - `ssl`
  - `os`
  - `time`
  - `random`

To install required dependencies, run:
```sh
pip install pandas
```

## Configuration
1. Update the following variables in `send_email` function:
   - `sender_email`: Your email address
   - `sender_password`: Your email password (consider using environment variables for security)
   - `smtp_server` and `smtp_port`: Update as per your email provider

2. Modify `excel_file_path` with the correct path to your Excel file containing recipient details.

3. Ensure that your Excel file contains the following columns:
   - `Reciever`: Recipient's name
   - `Reciever Email`: Recipient's email address
   - `User name`: Username for access
   - `Password`: Password for access

4. Update `attachments` list with the correct file paths for any attachments.

## Usage
Run the script with:
```sh
python email_sender.py
```

The script will:
- Read recipient details from the Excel file.
- Check the log file to avoid resending emails.
- Send emails with the provided credentials and attachments.
- Log successfully sent emails and errors.

## Error Handling
- Any errors encountered while sending emails are recorded in `error_log.csv`.
- Successfully sent emails are logged in `sent_emails.txt`.

## Security Considerations
- Avoid hardcoding credentials; use environment variables instead.
- Ensure that your email provider allows SMTP connections.
- Use app passwords if two-factor authentication (2FA) is enabled.

## Contribution
Feel free to contribute by:
- Improving error handling
- Enhancing security measures
- Adding new features
