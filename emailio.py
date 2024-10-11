import argparse
import csv
import datetime
import hashlib
import logging
import os
import sys
import uuid
from azure.cosmos import CosmosClient
from jinja2 import Template
from opencensus.ext.azure.log_exporter import AzureEventHandler
if sys.platform == 'win32':
    import win32com.client as win32
elif sys.platform == 'darwin':
    from appscript import app, k


def send_email(subject, row):

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


def log_email_sent(ai_connection_string, email_address, args, subject):
    ai_logger = logging.getLogger(__name__)
    ai_logger.addHandler(AzureEventHandler(connection_string=ai_connection_string))
    ai_logger.setLevel(logging.INFO)

    salted_email = email_address + "-foundrywebsite"
    email_hash = hashlib.sha256(salted_email.encode()).hexdigest()

    properties = {'custom_dimensions': {'email_hash': email_hash,
                                        'subject': subject,
                                        'campaign': args.campaign}}
    ai_logger.info('email-sent', extra=properties)


def update_cosmos(cosmos_endpoint,
                  cosmos_key,
                  email_address,
                  campaign):

    client = CosmosClient(cosmos_endpoint, cosmos_key)
    database = client.get_database_client("Website")
    container = database.get_container_client("Signups")

    query = "SELECT c.id FROM c WHERE c.email = @email"
    parameters = [
        {"name": "@email", "value": email_address}
    ]

    items = list(container.query_items(
        query=query,
        parameters=parameters,
        enable_cross_partition_query=True
    ))

    if len(items) > 1:
        raise Exception(f"Multiple items found for email {email_address}")

    if len(items) < 1:
        document_data = {
            'id': str(uuid.uuid4()),
            'email': email_address,
            'campaigns': [campaign],
            'timestamp8601': datetime.datetime.utcnow().isoformat()
        }
    else:
        item_id = items[0]["id"]
        document_data = container.read_item(item_id,
                                            partition_key=email_address)
        if 'campaigns' not in document_data:
            document_data['campaigns'] = []

        if campaign not in document_data['campaigns']:
            document_data['campaigns'].append(campaign)

    container.upsert_item(document_data)


parser = argparse.ArgumentParser(description="Send templated emails")

parser.add_argument("-t", "--template", required=True, help="Path to the email template file")
parser.add_argument("-d", "--data", required=True, help="Path to the data CSV file")
parser.add_argument("-s", "--subject", required=True, help="Email subject")
parser.add_argument("-c", "--campaign", required=True, help="Email campaign, gets logged in Cosmos")
parser.add_argument("-ac", "--ai-connection-string", required=True, help="Application Insights connection string")
parser.add_argument('-ce', '--cosmos-endpoint', type=str, required=True, help='Azure Cosmos DB endpoint.')
parser.add_argument('-ck', '--cosmos-key', type=str, required=True, help='Azure Cosmos DB key.')

args = parser.parse_args()

if not args.template:
    print("Error: Template file path must be provided.")
    sys.exit(1)

if not args.data:
    print("Error: Data file path must be provided.")
    sys.exit(2)

if not args.subject:
    print("Error: Email subject must be provided.")
    sys.exit(3)

if not args.campaign:
    print("Error: Email campaign must be provided.")
    sys.exit(4)

if not args.ai_connection_string:
    print("Error: Application Insights connection string must be provided.")
    sys.exit(5)

if not args.cosmos_endpoint:
    print("Error: Azure Cosmos DB endpoint must be provided.")
    sys.exit(6)

if not os.path.exists(args.template):
    print(f"Error: Template file '{args.template}' not found.")
    sys.exit(2)

if not os.path.exists(args.data):
    print(f"Error: Data file '{args.data}' not found.")
    sys.exit(3)

template_path = args.template
data_path = args.data
ai_connection_string = args.ai_connection_string
subject = args.subject

with open(template_path, 'r') as template_file:
    template_content = template_file.read()
    email_template = Template(template_content)

with open(data_path, mode='r') as csv_file:
    csv_reader = csv.DictReader(csv_file)

    print("Email sent to:")

    for row in csv_reader:

        send_email(subject, row)

        log_email_sent(ai_connection_string, row['email_address'], args, subject)

        update_cosmos(args.cosmos_endpoint,
                      args.cosmos_key,
                      row['email_address'],
                      args.campaign)

        print(row['email_address'])
