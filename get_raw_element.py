"""Get raw JSON for specific elements to see full style properties."""
import sys, os, json
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from google.oauth2 import service_account
from googleapiclient.discovery import build
from config import CREDENTIALS_FILE, SCOPES

TEMPLATE_ID = "1fAxAZUqg3LDF_CvRHnvjD7Zr0TbzVx4KRBwF4QVBT4g"
creds = service_account.Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
service = build("slides", "v1", credentials=creds)
pres = service.presentations().get(presentationId=TEMPLATE_ID).execute()

# Slide [3] P.4 — first two elements (ROUND_RECTANGLE + title)
slide = pres["slides"][3]
for elem in slide.get("pageElements", [])[:2]:
    print(json.dumps(elem, indent=2))
    print("---")
