import sys
from pathlib import Path
import xml.etree.ElementTree as ET

# Csak a standard lib + openpyxl használat, mivel xlsx-et írunk
try:
    from openpyxl import Workbook
except ImportError:
    print("Hiányzik az openpyxl csomag. Telepítsd: pip install openpyxl", file=sys.stderr)
    sys.exit(1)


def parse_ma2_xml(xml_path: Path):
    """
    Visszaad: (info_datetime, info_showfile, rows)
    rows: list[dict] a következő kulcsokkal:
      datetime, showfile, sequ_index, cue_index, cuepart_index, cuepart_name, basic_fade
    """
    # Namespace kezelés
    ns = {"ma": "http://schemas.malighting.de/grandma2/xml/MA"}

    tree = ET.parse(xml_path)
    root = tree.getroot()

    info = root.find("ma:Info", ns)
    info_datetime = info.attrib.get("datetime") if info is not None else ""
    info_showfile = info.attrib.get("showfile") if info is not None else ""

    rows = []
    # Végigmegyünk a Sequ elemeken
    for sequ in root.findall("ma:Sequ", ns):
        sequ_index = sequ.attrib.get("index")
        # Végigmegyünk a Cue elemeken
        for cue in sequ.findall("ma:Cue", ns):
            cue_index = cue.attrib.get("index")
            # Végigmegyünk a CuePart elemeken
            for cuepart in cue.findall("ma:CuePart", ns):
                rows.append({
                    "datetime": info_datetime,
                    "showfile": info_showfile,
                    "sequ_index": sequ_index,
                    "cue_index": cue_index,
                    "cuepart_index": cuepart.attrib.get("index"),
                    "cuepart_name": cuepart.attrib.get("name", ""),
                    "basic_fade": cuepart.attrib.get("basic_fade", "")
                })
    return info_datetime, info_showfile, rows


def write_xlsx(rows, out_path: Path):
    wb = Workbook()
    ws = wb.active
    ws.title = "cueparts"

    headers = ["datetime", "showfile", "sequ_index", "cue_index", "cuepart_index", "cuepart_name", "basic_fade"]
    ws.append(headers)

    for r in rows:
        ws.append([r.get(h, "") for h in headers])

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)


def main():
    # Bemeneti XML fájl: alapértelmezésben a projekt gyökerében lévő matyasexport.xml
    project_root = Path(__file__).resolve().parent
    xml_path = project_root / "matyasexport.xml"
    if len(sys.argv) > 1:
        xml_path = Path(sys.argv[1]).resolve()

    if not xml_path.exists():
        print(f"Nem található XML: {xml_path}", file=sys.stderr)
        sys.exit(2)

    _, _, rows = parse_ma2_xml(xml_path)

    if not rows:
        print("Nem találtam CuePart adatot az XML-ben.", file=sys.stderr)

    out_path = project_root / "xlsx" / f"{xml_path.stem}_cueparts.xlsx"
    write_xlsx(rows, out_path)

    print(f"Kész: {out_path}")


if __name__ == "__main__":
    main()