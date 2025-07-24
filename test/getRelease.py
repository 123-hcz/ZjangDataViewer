import requests
import json

r = requests.get('https://github.com/123-hcz/ZjangDataViewer/releases')

print(r.text)

