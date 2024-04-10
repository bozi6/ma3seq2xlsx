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
# Exported file(s) are in xls folder
#
import os
import platform
import xml.etree.ElementTree as ET

import openpyxl
from openpyxl.styles import Font
from openpyxl.styles.borders import Border, Side

if platform.system() == "Windows":
    default_ma3_seq_dir = (
        "C:\\ProgramData\\MALightingTechnology\\gma3_library\\datapools\\sequences\\"
    )
elif platform.system() == "Darwin":
    default_ma3_seq_dir = (
        "/Users/mnte/MALightingTechnology/gma3_library/datapools/sequences/"
    )

filenames = os.listdir(default_ma3_seq_dir)
menu = {}
menunum = 0
for filename in filenames:
    if filename.endswith(".xml"):
        menunum += 1
        menu[menunum] = filename

menu[len(menu) + 1] = "Exit"


def menuiras():
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
    menuiras()
    vfile = int(input("Choose xml file to convert to xlsxs."))
    if vfile > len(menu):
        print("Wrong number.")
        continue
    elif vfile == len(menu):
        exit(0)
    file = default_ma3_seq_dir + menu[vfile]
    print("Chosen XML file: ", file)
    root = ET.parse(file).getroot()
    verzio = root.attrib["DataVersion"]
    # seqnev = root.get("Sequence/Name")
    # print(seqnev)
    treeData = [["Num:", "Cue name:", "Note:", "Comment:"]]
    for type_tag in root.findall("Sequence/Cue"):
        try:
            sorszam = float(type_tag.get("No").strip())
        except TypeError:
            sorszam = 0
        except AttributeError:
            sorszam = 0
        nev = type_tag.get("Name")
        note = type_tag.get("Note")
        comment = ""
        treeData.append([sorszam, nev, note, comment])
        # print(sorszam, "--", nev)
    # print(treeData)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = menu[vfile][:-4]
    header = Font(size=24, italic=True)
    listtext = Font(size=16)
    ws["A1"].font = header
    ws["A1"] = "Seq name: " + ws.title
    ws["C1"].font = Font(size=16, bold=True)
    ws["C1"] = " MA3 program version: " + verzio
    ws["C2"].font = listtext
    ws.merge_cells(range_string="A1:B1")
    for tree in treeData:
        ws.append(tree)
    for row in ws.iter_rows(min_row=2, max_col=4):
        for cell in row:
            cell.border = thin_border
            cell.font = listtext
    ws.column_dimensions["A"].width = 11
    ws.column_dimensions["B"].width = 55
    ws.column_dimensions["C"].width = 55
    ws.column_dimensions["D"].width = 55
    try:
        wb.save(f"./xls/{ws.title}.xlsx")
    except FileNotFoundError:
        print("xls directory not found, so try to create.")
        os.makedirs("./xls")
        wb.save(f"./xls/{ws.title}.xlsx")
    wb.close()
    print("Writing file done.\n Restarting")
