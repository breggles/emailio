import win32com.client as win32
import csv
from jinja2 import Template
import sys
import os

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

        outlook = win32.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)
        mail.To = row['email_address']
        mail.Subject = row['subject']
        # mail.Body = row['plain_text_body']
        mail.HTMLBody = html_body

        mail.Send()
