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

emailio.py --template TEMPLATE_PATH --data DATA_PATH --ai-connection-string AI_CONNECTION_STRING --subject SUBJECT

## Data file format

The data file is in CSV format and has the following required fields:

* email_address - email address of the recipient

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
    <p>Company Inc.</p>
</body>
</html>
```
## Sample email data

```csv
email_address,name,description
joe.blogs@example.com,Joe Blogs,We are thrilled to announce the release of our newest product.
```

## Application Insights

Logs an `email-sent` custom event to AppInsights, with the `email-hash` and `subject` as a custom dimensions. The `email-hash` is a SHA256 hash of the email address and a salt.
