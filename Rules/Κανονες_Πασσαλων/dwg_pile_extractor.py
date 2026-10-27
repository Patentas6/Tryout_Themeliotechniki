# -*- coding: utf-8 -*-
"""
dwg_pile_extractor.py — Αυτόματη μετατροπή DWG -> DXF & εξαγωγή όλων των στοιχείων πασσάλων
(Κατάλογος Πασσάλων, Διάμετροι, Υψόμετρα, Σκυρόδεμα, Οπλισμός, Μάτιση, Αναμονές) σε 1 εντολή!

Χρήση:
    python dwg_pile_extractor.py "<path_to_dwg_or_dxf>"
"""
import os
import sys
import re
import json
import time

sys.stdout.reconfigure(encoding="utf-8")


def convert_dwg_to_dxf(dwg_path: str) -> str:
    dxf_path = os.path.splitext(dwg_path)[0] + "_extracted.dxf"
    if os.path.exists(dxf_path) and os.path.getmtime(dxf_path) >= os.path.getmtime(dwg_path):
        return dxf_path
    import win32com.client
    acad = None
    for prog_id in ("AutoCAD.Application", "AutoCAD.Application.25", "AutoCAD.Application.24"):
        try:
            acad = win32com.client.Dispatch(prog_id)
            break
        except Exception:
            continue
    if acad is None:
        raise RuntimeError("Δεν βρέθηκε εγκατεστημένο AutoCAD για μετατροπή DWG -> DXF.")
    doc = acad.Documents.Open(os.path.abspath(dwg_path), True)
    time.sleep(2)
    doc.SaveAs(os.path.abspath(dxf_path), 61)  # 61 = ac2013_dxf
    doc.Close(False)
    return dxf_path


def extract_pile_info(path: str) -> dict:
    ext = os.path.splitext(path)[1].lower()
    dxf_path = convert_dwg_to_dxf(path) if ext == ".dwg" else path

    import ezdxf
    doc = ezdxf.readfile(dxf_path)
    msp = doc.modelspace()

    texts = []
    for e in msp.query("TEXT MTEXT"):
        try:
            raw = e.plain_text() if e.dxftype() == "MTEXT" else e.dxf.text
            t = " ".join(raw.split()).strip()
            if not t:
                continue
            ins = e.dxf.insert
            texts.append({"text": t, "x": round(float(ins.x), 2), "y": round(float(ins.y), 2), "layer": e.dxf.layer})
        except Exception:
            pass

    kw_pat = re.compile(
        r"(πασσ|pile|ø|φ\d|y\d|c\d{2}/\d{2}|cut.?off|cast|cover|επικαλ|ματισ|lap|starter|αναμον|\bP\d{1,2}\b)",
        re.IGNORECASE,
    )
    relevant = [it for it in texts if kw_pat.search(it["text"])]
    relevant.sort(key=lambda r: (-r["y"], r["x"]))

    out_json = os.path.splitext(path)[0] + "_pile_extract.json"
    result = {
        "source_file": path,
        "dxf_file": dxf_path,
        "total_text_entities": len(texts),
        "relevant_count": len(relevant),
        "relevant_texts": relevant[:400],
    }
    with open(out_json, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    return result


if __name__ == "__main__":
    res = extract_pile_info(sys.argv[1])
    print(f"Extracted {res['relevant_count']} relevant pile texts (from {res['total_text_entities']} total).")
    for item in res["relevant_texts"][:120]:
        print(f"[{item['layer']}] ({item['x']}, {item['y']}): {item['text']}")
