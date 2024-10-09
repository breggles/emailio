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
from azure.cosmos import CosmosClient
import datetime
import uuid

parser = argparse.ArgumentParser(description="Send templated emails")

parser.add_argument("-t", "--template", required=True, help="Path to the email template file")
parser.add_argument("-d", "--data", required=True, help="Path to the data CSV file")
parser.add_argument("-s", "--subject", required=True, help="Email subject")
parser.add_argument("-c", "--campaign", required=True, help="Email campaign, gets logged in Cosmos")
parser.add_argument("-ac", "--ai-connection-string", required=True, help="Application Insights connection string")
parser.add_argument('-ce', '--cosmos-endpoint', type=str, required=True, help='Azure Cosmos DB endpoint.')
parser.add_argument('-ck', '--cosmos-key', type=str, required=True, help='Azure Cosmos DB key.')
parser.add_argument('-cn', '--cosmos-database', type=str, required=True, help='Azure Cosmos DB database name.')
parser.add_argument('-co', '--cosmos-container', type=str, required=True, help='Azure Cosmos DB container name.')

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

        # Cosmos

        client = CosmosClient(args.cosmos_endpoint, args.cosmos_key)
        database = client.get_database_client(args.cosmos_database)
        container = database.get_container_client(args.cosmos_container)

        query = "SELECT c.id FROM c WHERE c.email = @email"
        parameters = [
            {"name": "@email", "value": row['email_address']}
        ]

        items = list(container.query_items(
            query=query,
            parameters=parameters,
            enable_cross_partition_query=True
        ))

        if len(items) > 1:
            raise Exception(f"Multiple items found for email {row['email_address']}")

        if len(items) < 1:
            document_data = {
                'id': str(uuid.uuid4()),
                'email': row['email_address'],
                'campaigns': [args.campaign],
                'timestamp': datetime.datetime.utcnow().isoformat()
            }
        else:
            print(items[0]["id"])
            item_id = items[0]["id"]
            document_data = container.read_item(item_id,
                                                partition_key=item_id)
            print(document_data)
            document_data['campaigns'].append(args.campaign)

        container.upsert_item(document_data)
