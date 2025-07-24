import json

path = '../src/_json/options.json'

def load_json(file_path):
	with open(file_path, 'r', encoding='utf-8') as file:
		return json.load(file)

data = load_json(path)
print(data)
for i in data:
	print(i)
for i in data.values():
	print(i)
