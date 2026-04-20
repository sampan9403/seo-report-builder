import sys, os, json
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from google.oauth2 import service_account
from googleapiclient.discovery import build
from config import CREDENTIALS_FILE, SCOPES

TEMPLATE_ID = "1fAxAZUqg3LDF_CvRHnvjD7Zr0TbzVx4KRBwF4QVBT4g"
creds = service_account.Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
service = build("slides", "v1", credentials=creds)
pres = service.presentations().get(presentationId=TEMPLATE_ID).execute()

# 1. Get the slide layout used by P.4 (slide index 3)
slide4 = pres["slides"][3]
print("=== P.4 slideProperties ===")
print(json.dumps(slide4.get("slideProperties", {}), indent=2))

# 2. Get ELLIPSE element raw JSON (decorative dots) from P.4
print("\n=== P.4 ELLIPSE (first dot) ===")
group_elem = next(e for e in slide4.get("pageElements",[]) if "elementGroup" in e)
first_dot = group_elem["elementGroup"]["children"][0]
print(json.dumps(first_dot, indent=2))

# 3. Check P.10 slide layout
slide10 = pres["slides"][9]
print("\n=== P.10 slideProperties ===")
print(json.dumps(slide10.get("slideProperties",{}), indent=2))
