import requests, json, re
import sys
sys.stdout.reconfigure(encoding='utf-8')
r = requests.get('https://finance.vietstock.vn/search?query=TCB', headers={'User-Agent': 'Mozilla/5.0'})
print("Response text:", r.text)
