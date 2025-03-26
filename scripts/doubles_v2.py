
import os
import pprint
from azure.cosmos import CosmosClient
from collections import defaultdict

# Initialize the Cosmos client
cosmos_account_key = os.getenv("COSMOS_DB_KEY")

client = CosmosClient("https://redgate-innovation.documents.azure.com:443/", cosmos_account_key)

database = client.get_database_client("Website")

container = database.get_container_client("Signups")

# Query to get all documents with emails in lowercase
query = "SELECT * FROM c"
items = container.query_items(query=query, enable_cross_partition_query=True)

# Process to find duplicates
email_dict = defaultdict(list)
item_dict = {}
for item in items:
    email = item['email'].lower()
    email_dict[email].append(item['id'])
    item_dict[item['id']] = item

# Find and print duplicates
duplicates = {email: ids for email, ids in email_dict.items() if len(ids) > 1}
for email, ids in duplicates.items():
    print(f"\nDuplicate email: {email} found in documents:\n")
    for id in ids:
        pprint.pprint(item_dict[id])
