import json
print("cookies = ", end='')
vals = dict()
with open("cookies.json", "r") as f:
    for coc in json.load(f):
        vals[coc['name']] = coc['value']

print(json.dumps(vals))