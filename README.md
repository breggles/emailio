# Emailio

Sends templated emails through your Outlook. Like Word Mail Merge, but less shit. Also, it logs an event to AppInsights for every email sent. Also updates the user in a Cosmos DB database.

## Supported OSs

Tested on Windows. Might work on macOS...

## Prerequisites

Run this to install required packages:

```bash
pip install pywin32 appscript jinja2 opencensus-ext-azure
```

## Usage

```sh
emailio.py \
    --template TEMPLATE_PATH \
    --data DATA_PATH -\
    --subject SUBJECT
    --campaign CAMPAIGN \
    --ai-connection-string AI_CONNECTION_STRING -\
    --cosmos-endpoint COSMOS_ENDPOINT \
    --cosmos-key COSMOS_KEY
```

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

    <p>We are excited to inform you about this that and the other.</p>

    <p><a href="{{ url }}">This is a link</a>.</p>

    <p>Best regards,</p>
    <p>Company Inc.</p>
</body>
</html>
```
## Sample email data

```csv
email_address,name,url
joe.blogs@example.com,Joe Blogs,https://example.com
```

## Application Insights

For every email sent, logs an `email-sent` custom event to AppInsights, with the `email-hash`, `campaign` and `subject` as a custom dimensions. The `email-hash` is a SHA256 hash of the email address and a salt.

## Cosmos DB

For every email sent, updates user in the Signups container of the Website database. Adds the `campaign` to the `campaigns` array.
