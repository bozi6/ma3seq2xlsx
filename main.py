#  main.py Copyright (C) 2024  Konta Boáz
#      This program comes with ABSOLUTELY NO WARRANTY;.
#      This is free software, and you are welcome to redistribute it
#      under certain conditions;
#   Last Modified: 2024. 04. 10. 19:51
# Usage:
# In ma3 console/onPC
#
# in command line enter:
# cd sequence
# list
# export xx "namewhatyouwant.xml"
# cd root
# exported sequences are in
# MacOS: [System HD]/Users/[User Name]/MALightingTechnology/gma3_library/datapools/sequences/
# Windows: C:\ProgramData\MALightingTechnology\gma3_library\datapools\sequences/
#
# Exported file(s) are in xlsx folder
#
import os
import platform
import xml.etree.ElementTree as et

import openpyxl
from openpyxl.styles import Font
from openpyxl.styles.borders import Border, Side

default_ma3_seq_dir = ""

if platform.system() == "Windows":
    default_ma3_seq_dir = (
        "C:\\ProgramData\\MALightingTechnology\\gma3_library\\datapools\\sequences\\"
    )
elif platform.system() == "Darwin":
    homedir = os.path.expanduser("~")
    default_ma3_seq_dir = (
        f'{homedir}/MALightingTechnology/gma3_library/datapools/sequences/'3
    )

filenames = os.listdir(default_ma3_seq_dir)
menu = {}
menunum = 0
for filename in filenames:
    if filename.endswith(".xml"):
        menunum += 1
        menu[menunum] = filename

menu[len(menu) + 1] = "Exit"


def menushow():
    print("-" * 30)
    print("MENU")
    print("-" * 30)

    for key in menu.keys():
        print(key, "--", menu[key])


thin_border = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)

while True:
    menushow()
    vfile = int(input("Choose xml file to convert to xlsxs."))
    if vfile > len(menu):
        print("Wrong number.")
        continue
    elif vfile == len(menu):
        exit(0)
    file = default_ma3_seq_dir + menu[vfile]
    print("XML file: ", file)
    root = et.parse(file).getroot()
    verzio = root.attrib["DataVersion"]
    try:
        seqnote = root.find(".//Sequence").attrib.get("Note").replace("&#xD;", " ")
        seqnote = seqnote.replace("\r", " ")
        seqnote = seqnote.strip("  ")
    except AttributeError:
        seqnote = ""
    treeData = [
        ["Num:", "Cue name:", "FadeIn:", "FadeOut", "Cue Note:", "Trig.Type/Param:", "Comment:"]
    ]
    for type_tag in root.iter("Cue"):
        try:
            sorszam = float(type_tag.get("No").strip())
        except TypeError:
            sorszam = 0
        except AttributeError:
            sorszam = 0
        nev = type_tag.get("Name")
        note = type_tag.get("Note")
        cuefadein = ''
        cuefadeout = ''
        for child in type_tag:
            if "CueInFade" in child.attrib:
                cuefadein = child.attrib['CueInFade']
            if "CueOutFade" in child.attrib:
                cuefadeout = child.attrib['CueOutFade']
        trigtype = type_tag.get("TrigType")
        if trigtype is None:
            trigtype = "Go+"
        elif trigtype == "Time":
            ido = type_tag.get("TrigTime")
            trigtype = f"{trigtype} - {ido}"
        elif trigtype == "Sound":
            hang = type_tag.get("TrigSound")
            trigtype = f"{trigtype} - {hang}"
        comment = ""
        treeData.append([sorszam, nev, cuefadein, cuefadeout, note, trigtype, comment])

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = menu[vfile][:-4]
    header = Font(size=24, italic=True)
    listtext = Font(size=16)
    vertext = Font(size=14, bold=True)
    notetext = Font(size=14, italic=True)
    ws["A1"].font = header
    ws["A1"] = "Seqence name: " + ws.title
    ws["A2"].font = vertext
    ws["A2"] = " MA3 program version: " + verzio
    ws["A3"].font = header
    ws["A3"] = "Sequence note: "
    ws["A4"].font = notetext
    ws["A4"] = seqnote
    ws.merge_cells(range_string="A1:D1")
    for tree in treeData:
        ws.append(tree)
    for row in ws.iter_rows(min_row=5, max_col=7):
        for cell in row:
            cell.border = thin_border
            cell.font = listtext
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 55
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 55
    ws.column_dimensions["F"].width = 20
    ws.column_dimensions["G"].width = 55
    try:
        wb.save(f"./xlsx/{ws.title}.xlsx")
    except FileNotFoundError:
        print("xlsx directory not found, try to create it.")
        os.makedirs("./xlsx")
        wb.save(f"./xlsx/{ws.title}.xlsx")
    wb.close()
    print("Writing file done.\n Restarting")
