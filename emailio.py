import sys
import os
if sys.platform == 'win32':
    import win32com.client as win32
elif sys.platform == 'darwin':
    from appscript import app, k
import csv
from jinja2 import Template
import logging
from opencensus.ext.azure.log_exporter import AzureEventHandler
import argparse
import hashlib

parser = argparse.ArgumentParser(description="Send templated emails")
parser.add_argument("-t", "--template", required=True, help="Path to the email template file")
parser.add_argument("-d", "--data", required=True, help="Path to the data CSV file")
parser.add_argument("-c", "--ai-connection-string", required=True, help="Application Insights connection string")
parser.add_argument("-s", "--subject", required=True, help="Email subject")
args = parser.parse_args()

template_path = args.template
data_path = args.data
ai_connection_string = args.ai_connection_string
subject = args.subject

if not os.path.exists(template_path):
    print(f"Error: Template file '{template_path}' not found.")
    sys.exit(2)

if not os.path.exists(data_path):
    print(f"Error: Data file '{data_path}' not found.")
    sys.exit(3)

if not ai_connection_string:
    print("Error: Application Insights connection string must be provided as a command line argument.")
    sys.exit(4)

if not subject:
    print("Error: Email subject must be provided as a command-line argument.")
    sys.exit(5)

ai_logger = logging.getLogger(__name__)
ai_logger.addHandler(AzureEventHandler(connection_string=ai_connection_string))
ai_logger.setLevel(logging.INFO)

with open(template_path, 'r') as template_file:
    template_content = template_file.read()
    email_template = Template(template_content)

with open(data_path, mode='r') as csv_file:
    csv_reader = csv.DictReader(csv_file)

    for row in csv_reader:
        html_body = email_template.render(row)

        if sys.platform == 'win32':
            outlook = win32.Dispatch('Outlook.Application')
            mail = outlook.CreateItem(0)
            mail.To = row['email_address']
            mail.Subject = subject
            mail.HTMLBody = html_body
            mail.Send()

        elif sys.platform == 'darwin':
            recipient = app('Mail').make(new=k.to_recipient,
                                         with_properties={k.address: row['email_address']})
            mail = app('Mail').make(new=k.outgoing_message)
            mail.subject.set(subject)
            mail.content.set(html_body)
            mail.to_recipients.set([recipient])
            mail.send()

        salted_email = row["email_address"] + "-foundrywebsite"

        email_hash = hashlib.sha256(salted_email.encode()).hexdigest()

        properties = {'custom_dimensions': {'email_hash': email_hash,
                                            'subject': subject}}
        ai_logger.info('email-sent', extra=properties)
