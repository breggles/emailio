from azure.cosmos import CosmosClient

cosmos_endpoint = "https://redgate-innovation.documents.azure.com:443/"
cosmos_key = ""

client = CosmosClient(cosmos_endpoint, cosmos_key)
database = client.get_database_client("Website")
container = database.get_container_client("Signups")

query = "SELECT * FROM c WHERE IS_DEFINED(c.timestamp)"

items = list(container.query_items(
    query=query,
    enable_cross_partition_query=True
))

for item in items:

    print(item["id"])

    doc = container.read_item(item["id"], partition_key=item["email"])

    doc['timestamp8601'] = doc.pop('timestamp')

    container.delete_item(item["id"], partition_key=item["email"])

    container.upsert_item(doc)
