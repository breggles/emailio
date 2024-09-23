# Emailio

Sends templated emails through your Outlook. Like Word Mail Merge, but less shit.

## Supported OSs

Tested on Windows. Might work on macOS...

## Prerequisites

Run this to install required packages:

```bash
pip install pywin32 appscript jinja2 opencensus-ext-azure
```

## Usage

emailio.py email_template.html email_data.csv

## Special fields

* email_address - email address of the recipient
* subject - subject of the email

## Sample email template

```html
<!DOCTYPE html>
<html>
<head>
    <title>Email</title>
</head>
<body>
    <p>Dear {{ name }},</p>

    <p>We are excited to inform you about {{ subject }}.</p>

    <p>{{ description }}</p>

    <p>Best regards,</p>
    <p>{{ sender_name }}</p>
</body>
</html>
```
## Sample email data

```csv
name,email,subject,description,sender_name
Joe Blogs,joe.blogs@example.com,New Product Launch,We are thrilled to announce the release of our newest product.,Company Inc.
```

## Application Insights

If the `APPLICATIONINSIGHTS_CONNECTION_STRING` environment variable is set, logs an `email-sent` event to AppInsights, with the `subject` as a custom dimension.

NB: requires a `email_hash` field in the csv data, which is used to populate the `email_hash` custom dimension.
