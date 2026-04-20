"""Inspect Slide 10 & 11 elements to see actual text content."""
import sys, os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from google.oauth2 import service_account
from googleapiclient.discovery import build
from config import CREDENTIALS_FILE, SCOPES

PRES_ID = input("Paste presentation ID: ").strip()

creds = service_account.Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
service = build("slides", "v1", credentials=creds)
pres = service.presentations().get(presentationId=PRES_ID).execute()
slides = pres["slides"]

for slide_idx, label in [(9, "SLIDE 10"), (10, "SLIDE 11")]:
    slide = slides[slide_idx]
    print(f"\n{'='*60}")
    print(f"{label} — {len(slide.get('pageElements', []))} elements")
    print('='*60)
    for i, elem in enumerate(slide.get("pageElements", [])):
        elem_id = elem.get("objectId", "?")
        if "shape" in elem:
            shape = elem["shape"]
            text_elems = shape.get("text", {}).get("textElements", [])
            text = "".join(
                te.get("textRun", {}).get("content", "")
                for te in text_elems
            ).strip()
            shape_type = shape.get("shapeType", "UNKNOWN")
            print(f"  [{i}] SHAPE ({shape_type}) id={elem_id}")
            if text:
                print(f"       TEXT: {repr(text[:120])}")
            else:
                print(f"       TEXT: (empty)")
        elif "image" in elem:
            size = elem.get("size", {})
            w = size.get("width", {}).get("magnitude", 0) / 12700
            h = size.get("height", {}).get("magnitude", 0) / 12700
            print(f"  [{i}] IMAGE id={elem_id} size={w:.0f}x{h:.0f}pt")
        elif "table" in elem:
            print(f"  [{i}] TABLE id={elem_id}")
        else:
            keys = list(elem.keys())
            print(f"  [{i}] OTHER {keys} id={elem_id}")

