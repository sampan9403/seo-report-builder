"""Extract exact positions, sizes, and styles from reference template Slide 10 & 11."""
import sys, os, json
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from google.oauth2 import service_account
from googleapiclient.discovery import build
from config import CREDENTIALS_FILE, SCOPES

TEMPLATE_ID = "1fAxAZUqg3LDF_CvRHnvjD7Zr0TbzVx4KRBwF4QVBT4g"

creds = service_account.Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
service = build("slides", "v1", credentials=creds)
pres = service.presentations().get(presentationId=TEMPLATE_ID).execute()
slides = pres["slides"]

def emu_to_pt(emu):
    return round(emu / 12700, 2)

def get_transform(elem):
    t = elem.get("transform", {})
    size = elem.get("size", {})
    tx = t.get("translateX", 0)
    ty = t.get("translateY", 0)
    sx = t.get("scaleX", 1)
    sy = t.get("scaleY", 1)
    w = size.get("width", {}).get("magnitude", 0)
    h = size.get("height", {}).get("magnitude", 0)
    return {
        "x_pt":  emu_to_pt(tx),
        "y_pt":  emu_to_pt(ty),
        "w_pt":  emu_to_pt(w * sx),
        "h_pt":  emu_to_pt(h * sy),
        "x_emu": tx, "y_emu": ty,
        "w_emu": w,  "h_emu": h,
        "scaleX": sx, "scaleY": sy,
    }

def get_text_runs(elem):
    shape = elem.get("shape", {})
    runs = []
    for te in shape.get("text", {}).get("textElements", []):
        if "textRun" in te:
            content = te["textRun"].get("content", "")
            style = te["textRun"].get("style", {})
            fg = style.get("foregroundColor", {}).get("opaqueColor", {}).get("rgbColor", {})
            runs.append({
                "text": repr(content),
                "bold": style.get("bold", False),
                "font": style.get("fontFamily", ""),
                "size_pt": style.get("fontSize", {}).get("magnitude", 0),
                "color_rgb": fg,
            })
    return runs

def print_elem(elem, indent=""):
    t = get_transform(elem)
    obj_id = elem.get("objectId", "?")

    if "image" in elem:
        print(f"{indent}IMAGE  id={obj_id}")
        print(f"{indent}  pos: x={t['x_pt']}pt  y={t['y_pt']}pt")
        print(f"{indent}  size: w={t['w_pt']}pt  h={t['h_pt']}pt")
        print(f"{indent}  EMU: x={t['x_emu']}  y={t['y_emu']}  w={t['w_emu']}  h={t['h_emu']}")

    elif "shape" in elem:
        shape = elem["shape"]
        shape_type = shape.get("shapeType", "?")
        runs = get_text_runs(elem)
        full_text = "".join(r["text"] for r in runs)
        print(f"{indent}SHAPE ({shape_type})  id={obj_id}")
        print(f"{indent}  pos: x={t['x_pt']}pt  y={t['y_pt']}pt")
        print(f"{indent}  size: w={t['w_pt']}pt  h={t['h_pt']}pt")
        print(f"{indent}  EMU: x={t['x_emu']}  y={t['y_emu']}  w={t['w_emu']}  h={t['h_emu']}")
        if runs:
            print(f"{indent}  text (first run): {full_text[:80]}")
            for r in runs[:3]:
                print(f"{indent}    run: bold={r['bold']} font={r['font']!r} size={r['size_pt']}pt color={r['color_rgb']}")

    elif "elementGroup" in elem:
        children = elem.get("elementGroup", {}).get("children", [])
        print(f"{indent}GROUP  id={obj_id}  ({len(children)} children)")
        for c in children:
            print_elem(c, indent + "  ")

for slide_idx, label in [(9, "SLIDE 10"), (10, "SLIDE 11")]:
    slide = slides[slide_idx]
    print(f"\n{'='*70}")
    print(f"{label}")
    print(f"Slide size: {emu_to_pt(pres['pageSize']['width']['magnitude'])}pt x {emu_to_pt(pres['pageSize']['height']['magnitude'])}pt")
    print('='*70)
    for elem in slide.get("pageElements", []):
        print_elem(elem)
        print()
