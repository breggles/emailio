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

ai_connection_string = os.getenv('APPLICATIONINSIGHTS_CONNECTION_STRING')

if not ai_connection_string:
    print("Application Insights connection string not provided.")
    exit(1)

ai_logger = logging.getLogger(__name__)
ai_logger.addHandler(AzureEventHandler(connection_string=ai_connection_string))
ai_logger.setLevel(logging.INFO)

if len(sys.argv) < 3:
    print("Error: Template file and data file must be provided as arguments.")
    sys.exit(3)

template_path = sys.argv[1]
data_path = sys.argv[2]

if not os.path.exists(template_path):
    print(f"Error: Template file '{template_path}' not found.")
    sys.exit(1)

if not os.path.exists(data_path):
    print(f"Error: Data file '{data_path}' not found.")
    sys.exit(2)

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
            mail.Subject = row['subject']
            mail.HTMLBody = html_body
            mail.Send()

        elif sys.platform == 'darwin':
            recipient = app('Mail').make(new=k.to_recipient,
                                         with_properties={k.address: row['email_address']})
            mail = app('Mail').make(new=k.outgoing_message)
            mail.subject.set(row['subject'])
            mail.content.set(html_body)
            mail.to_recipients.set([recipient])
            mail.send()

        properties = {'custom_dimensions': {'email_hash': row['email_hash'],
                                            'subject': row['subject']}}
        ai_logger.info('email-sent', extra=properties)
