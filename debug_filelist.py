import requests
from main import get_token, get_file_list, parse_file_items

session = requests.Session()
try:
    token = get_token(session)
except Exception as e:
    print("get_token failed:", e)
    raise

case_no = "113108021"  # sample failing case from download_log.csv
print("Calling get_file_list for case_no:", case_no)
fl = get_file_list(session, token, case_no)
print("Raw response type:", type(fl))
# print a short preview
import json
try:
    print(json.dumps(fl, ensure_ascii=False)[:2000])
except Exception:
    print(str(fl)[:2000])

items = parse_file_items(fl)
print("Parsed items:")
for i, (name, fid) in enumerate(items, 1):
    print(i, name, fid)
