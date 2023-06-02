import smtplib
import openpyxl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# Email configuration
sender_email = 'abc@abc.com'  # Replace with your own email address
sender_password = 'abcdefghijklmnop'  # Replace with your own email password
subject = 'Your subject here'
body = '''Your body. In multiple lines. Here
'''

# Excel file configuration
excel_file = 'email.xlsx'  # Replace with your Excel file name and location
sheet_name = 'Sheet1'  # Replace with your sheet name
email_column = 'A'  # Replace with the column letter where email addresses are stored

# Attachment configuration
attachment_file = 'abc.txt'  # Replace with your attachment file name

# Read email addresses from Excel file
wb = openpyxl.load_workbook(excel_file)
sheet = wb[sheet_name]
email_addresses = [cell.value for cell in sheet[email_column] if cell.value]

# Connect to SMTP server
smtp_server = 'smtp.gmail.com'  # Replace with your SMTP server address
smtp_port = 587  # Replace with your SMTP server port
smtp_connection = smtplib.SMTP(smtp_server, smtp_port)
smtp_connection.starttls()
smtp_connection.login(sender_email, sender_password)

for email_address in email_addresses:
    # Create email message
    message = MIMEMultipart()
    message['From'] = sender_email
    message['To'] = email_address
    message['Subject'] = subject

    # Add body text to email
    message.attach(MIMEText(body, 'plain'))

    # Add attachment
    with open(attachment_file, 'rb') as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename= {attachment_file}')
        message.attach(part)

    # Send email
    smtp_connection.send_message(message)
    print(f"Email sent to {email_address}")

# Close the SMTP connection
smtp_connection.quit()
