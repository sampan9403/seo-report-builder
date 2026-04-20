"""Debug: show all elements on slide 10 & 11 and test matchers."""
import sys, os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from google.oauth2 import service_account
from googleapiclient.discovery import build
from config import CREDENTIALS_FILE, SCOPES

raw = input("Paste the SAME presentation ID or URL you used in the app: ").strip()
if "presentation/d/" in raw:
    PRES_ID = raw.split("presentation/d/")[1].split("/")[0]
else:
    PRES_ID = raw
print(f"Using ID: {PRES_ID}")

creds = service_account.Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
service = build("slides", "v1", credentials=creds)
pres = service.presentations().get(presentationId=PRES_ID).execute()
slides = pres["slides"]

def get_text_content(elem):
    shape = elem.get("shape", {})
    return " ".join(
        te.get("textRun", {}).get("content", "")
        for te in shape.get("text", {}).get("textElements", [])
    )

def print_elements(elements, indent=0):
    prefix = "  " * indent
    for i, elem in enumerate(elements):
        elem_id = elem.get("objectId", "?")
        keys = [k for k in elem.keys() if k not in ("objectId","transform","size")]
        elem_type = keys[0] if keys else "?"
        text = get_text_content(elem).strip()

        has_image = "image" in elem
        has_ranking = "ranking of" in text.lower()
        has_since = ("since the" in text.lower() or
                     "project started" in text.lower() or
                     "keywords reached" in text.lower())

        flags = []
        if has_image:   flags.append("★IMAGE")
        if has_ranking: flags.append("★RANKING_OF")
        if has_since:   flags.append("★SINCE_THE")

        flag_str = "  " + " ".join(flags) if flags else ""
        print(f"{prefix}[{i}] {elem_type.upper()} id={elem_id}{flag_str}")
        if text:
            print(f"{prefix}     → {repr(text[:100])}")

        # Recurse into elementGroup children
        children = elem.get("elementGroup", {}).get("children", [])
        if children:
            print(f"{prefix}     └─ GROUP children ({len(children)}):")
            print_elements(children, indent + 2)

for slide_idx, label in [(9, "SLIDE 10"), (10, "SLIDE 11")]:
    slide = slides[slide_idx]
    print(f"\n{'='*60}")
    print(f"{label} — {len(slide.get('pageElements', []))} top-level elements")
    print('='*60)
    print_elements(slide.get("pageElements", []))
