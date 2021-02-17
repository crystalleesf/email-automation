from googleapiclient.discovery import build
from google.oauth2 import service_account
from string import Template
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import json
import getpass


# Set up Google API credentials
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
SERVICE_ACCOUNT_FILE = './credentials.json'

credentials = None
credentials = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)

with open("./email_config.json", "r") as f:
    email_config = json.load(f)

SAMPLE_SPREADSHEET_ID = email_config['sheet_id']
HTML_MSG = email_config['html_msg']
PLAIN_MSG = email_config['plain_msg']
SUBJECT = email_config['subject']



def read_sheets():
    """
    Reads from Google Sheets
    """
    service = build('sheets', 'v4', credentials=credentials)
    # Call the Sheets API
    sheet = service.spreadsheets()
    # !!! ALWAYS REMEMBER TO UPDATE THE APPROPRIATE RANGE !!!
    result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range="Sheet1!A1:C4").execute()
    values = result.get('values', [])
    return values

def read_message_template(filename):
    with open(filename, 'r', encoding='utf-8') as tf:
        template_content = tf.read()
    return Template(template_content)

def send_email(contacts):
    MY_EMAIL = input("Enter email: ")
    PASSWORD = getpass.getpass("Enter password: ")

    with smtplib.SMTP_SSL(host='smtp.gmail.com', port=465) as server:
        server.login(MY_EMAIL, PASSWORD)
        # !!! ALWAYS REMEMBER TO UPDATE THE APPROPRIATE RANGE !!!
        for count, contact in enumerate(contacts, start=1):
            try:
                print("Sending Email {} ...".format(count))
                contact_first_name, contact_last_name, contact_email = contact[1].split()
                contact_email = contact_email[1:-1]

                message = MIMEMultipart('alternative')
                message_template = read_message_template(PLAIN_MSG)
                content_plain = message_template.substitute(CONTACT_NAME=contact_first_name, COMPANY=contact[0])
                with open(HTML_MSG, "r") as f:
                    content_html = f.read().format(CONTACT_NAME=contact_first_name, COMPANY=contact[0])
                
                message['From']="Crystal Lee <{}>".format(MY_EMAIL)
                message['To']="{} {}<{}>".format(contact_first_name, contact_last_name, contact_email)
                message['Subject']=SUBJECT


                # Adds image to email signature, but sometimes shows up as an attachment
                # with open('./hack.png', 'rb') as fp:
                #     msg_image = MIMEImage(fp.read())

                # Define the image's ID as referenced above
                # msg_image.add_header('Content-ID', '<image1>')
                # message.attach(msg_image)

                # Server will try to the html first, then the plain text version
                message.attach(MIMEText(content_plain, 'plain'))
                message.attach(MIMEText(content_html, 'html'))

                server.sendmail(MY_EMAIL, contact_email, message.as_string())
            
            except Exception as e:
                print(e)
                print("Email to {} failed".format(contact_email))

    print("All emails sent")



if __name__ == '__main__':
    contacts = read_sheets()
    send_email(contacts)
