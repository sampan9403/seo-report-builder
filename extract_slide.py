"""Extract one specific slide's full layout."""
import sys, os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from google.oauth2 import service_account
from googleapiclient.discovery import build
from config import CREDENTIALS_FILE, SCOPES

TEMPLATE_ID = "1fAxAZUqg3LDF_CvRHnvjD7Zr0TbzVx4KRBwF4QVBT4g"
SLIDE_INDEX = int(sys.argv[1]) if len(sys.argv) > 1 else 3

creds = service_account.Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
service = build("slides", "v1", credentials=creds)
pres = service.presentations().get(presentationId=TEMPLATE_ID).execute()
slide = pres["slides"][SLIDE_INDEX]

def emu(v): return round(v/12700, 1)

def dump(elem, indent=""):
    eid = elem.get("objectId","?")
    t = elem.get("transform",{}); sz = elem.get("size",{})
    tx=t.get("translateX",0); ty=t.get("translateY",0)
    sx=t.get("scaleX",1); sy=t.get("scaleY",1)
    w=sz.get("width",{}).get("magnitude",0); h=sz.get("height",{}).get("magnitude",0)

    if "image" in elem:
        print(f"{indent}IMAGE id={eid}")
        print(f"{indent}  x={emu(tx)}pt y={emu(ty)}pt  w={emu(w*sx)}pt h={emu(h*sy)}pt")
        print(f"{indent}  x_emu={int(tx)} y_emu={int(ty)} w_emu={int(w)} h_emu={int(h)}")
    elif "shape" in elem:
        shape = elem["shape"]
        stype = shape.get("shapeType","?")
        fill = shape.get("shapeProperties",{}).get("shapeBackgroundFill",{})
        bg = fill.get("solidFill",{}).get("color",{}).get("rgbColor",{})
        print(f"{indent}SHAPE({stype}) id={eid}")
        print(f"{indent}  x={emu(tx)}pt y={emu(ty)}pt  w={emu(w*sx)}pt h={emu(h*sy)}pt")
        print(f"{indent}  x_emu={int(tx)} y_emu={int(ty)}")
        if bg: print(f"{indent}  bg_color={bg}")
        for te in shape.get("text",{}).get("textElements",[]):
            if "textRun" in te:
                content = te["textRun"].get("content","")
                style = te["textRun"].get("style",{})
                fg = style.get("foregroundColor",{}).get("opaqueColor",{}).get("rgbColor",{})
                bold = style.get("bold",False)
                font = style.get("fontFamily","")
                size = style.get("fontSize",{}).get("magnitude",0)
                print(f"{indent}  run: {repr(content[:60])}  bold={bold} font={font!r} size={size}pt color={fg}")
            elif "paragraphMarker" in te:
                ps = te["paragraphMarker"].get("style",{})
                align = ps.get("alignment","")
                if align: print(f"{indent}  paragraph: align={align}")
    elif "elementGroup" in elem:
        children = elem.get("elementGroup",{}).get("children",[])
        print(f"{indent}GROUP id={eid} ({len(children)} children)")
        for c in children:
            dump(c, indent+"  ")

print(f"\nSlide [{SLIDE_INDEX}]  {len(slide.get('pageElements',[]))} top-level elements\n")
for e in slide.get("pageElements",[]):
    dump(e)
    print()
