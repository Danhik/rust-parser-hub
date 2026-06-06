import urllib.request, json

payload = json.dumps({
    'source_file': 'avan_my_inventory.csv',
    'dest_file':   'lis_catalog.csv',
    'min_profit':  1.0,
}).encode()

req = urllib.request.Request(
    'http://localhost:5050/api/build-links',
    data=payload,
    headers={'Content-Type': 'application/json'}
)
r    = urllib.request.urlopen(req)
data = json.load(r)

print(f"Matched: {data['matched']} | Source: {data['source_total']} | Dest: {data['dest_total']}")
for item in data['items'][:5]:
    print(f"  {item['name']}: buy {item['source_price']} -> sell {item['dest_price']} | +{item['profit']} ({item['roi']}%)")
