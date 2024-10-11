from azure.cosmos import CosmosClient

cosmos_endpoint = "https://redgate-innovation.documents.azure.com:443/"
cosmos_key = ""

client = CosmosClient(cosmos_endpoint, cosmos_key)
database = client.get_database_client("Website")
container = database.get_container_client("Signups")

email_addresses = ['se.vanvliet@apeldoorn.nl',
                   'kpatel@goldenstatefoods.com',
                   'dtai@friscotexas.gov',
                   'christian.moller@zisson.com',
                   'andy.aelbrecht.external@arcelormittal.com']

for email_address in email_addresses:

    print(f"Processing {email_address}")

    query = "SELECT c.id FROM c WHERE c.email = @email"
    parameters = [
        {"name": "@email", "value": email_address}
    ]

    items = list(container.query_items(
        query=query,
        parameters=parameters,
        enable_cross_partition_query=True
    ))

    doc0 = container.read_item(items[0]["id"],
                                        partition_key=email_address)

    doc1 = container.read_item(items[1]["id"],
                                        partition_key=email_address)

    if doc0['campaigns'] == None:
        source = doc1
        target = doc0
    else:
        source = doc0
        target = doc1

    target['campaigns'] = source['campaigns']

    container.upsert_item(target)

    container.delete_item(source['id'], partition_key=email_address)
