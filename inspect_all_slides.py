"""Inspect ALL slides in a presentation — shows slide index, title text, element count."""
import sys, os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from google.oauth2 import service_account
from googleapiclient.discovery import build
from config import CREDENTIALS_FILE, SCOPES

raw = input("Presentation URL or ID: ").strip()
if "presentation/d/" in raw:
    PRES_ID = raw.split("presentation/d/")[1].split("/")[0]
else:
    PRES_ID = raw

creds = service_account.Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
service = build("slides", "v1", credentials=creds)
pres = service.presentations().get(presentationId=PRES_ID).execute()
slides = pres["slides"]
print(f"\nPresentation: {pres.get('title','?')}  ({len(slides)} slides)\n")

def all_text(slide):
    texts = []
    def collect(elems):
        for e in elems:
            for te in e.get("shape",{}).get("text",{}).get("textElements",[]):
                t = te.get("textRun",{}).get("content","").strip()
                if t: texts.append(t)
            collect(e.get("elementGroup",{}).get("children",[]))
    collect(slide.get("pageElements",[]))
    return " | ".join(texts)

for i, slide in enumerate(slides):
    slide_id = slide["objectId"]
    n_elem = len(slide.get("pageElements",[]))
    text_preview = all_text(slide)[:120]
    print(f"  [{i}] id={slide_id}  elements={n_elem}")
    print(f"       {text_preview}")
    print()
