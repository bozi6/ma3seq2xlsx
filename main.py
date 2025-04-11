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

if platform.system() == "Windows":
    DEFAULT_MA3_SEQ_DIR = (
        "C:\\ProgramData\\MALightingTechnology\\gma3_library\\datapools\\sequences\\"
    )
elif platform.system() == "Darwin":
    homedir = os.path.expanduser("~")
    DEFAULT_MA3_SEQ_DIR = (
        f'{homedir}/MALightingTechnology/gma3_library/datapools/sequences/'
    )
else:
    DEFAULT_MA3_SEQ_DIR = ""

if not os.path.isdir(DEFAULT_MA3_SEQ_DIR):
    print(f"Error: Directory '{DEFAULT_MA3_SEQ_DIR}' does not exist or is inaccessible.")
    exit(1)

filenames = os.listdir(DEFAULT_MA3_SEQ_DIR)
menu = {}
menunum = 0
for filename in filenames:
    if filename.endswith(".xml"):
        menunum += 1
        menu[menunum] = filename

menu[len(menu) + 1] = "Exit"


def menushow():
    """
    Displays a formatted menu to the console.

    The function prints a menu header, followed by a list of menu items
    and their corresponding descriptions. The menu items are retrieved
    from the global `menu` dictionary, where the keys represent the
    menu options and the values represent their descriptions.

    Note:
        Ensure that the `menu` dictionary is defined and populated
        before calling this function to avoid errors.
    """
    print("-" * 30)
    print("MENU")
    print("-" * 30)

    for key in menu.keys():
        print(key, "--", menu[key])


thin_border = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"))

try:
    vfile = int(input("Choose xml file to convert to xlsxs."))
except ValueError:
    print("Invalid input. Please enter a number.")

running = True
while running:
    menushow()
    try:
        vfile = int(input("Choose xml file to convert to xlsxs."))
        if vfile > len(menu):
            print("Invalid selection. Please choose a valid menu option.")
            continue
        if vfile == len(menu):
            running = False
            break
    except ValueError:
        print("Invalid input. Please enter a number.")
        continue
    file = DEFAULT_MA3_SEQ_DIR + menu[vfile]
    print("XML file: ", file)
    root = et.parse(file).getroot()
    verzio = root.attrib["DataVersion"]
    seq_element = root.find(".//Sequence")
    if seq_element is not None and "Note" in seq_element.attrib:
        seqnote = seq_element.attrib["Note"].replace("&#xD;", " ")
        seqnote = seqnote.replace("\r", " ")
        seqnote = seqnote.strip("  ")
    else:
        seqnote = ""
    treeData = [
        ["Num:", "Cue name:", "FadeIn:", "FadeOut", "Cue Note:", "Trig.Type/Param:", "Comment:"]
    ]
    def process_cue(cue):
        try:
            sorszam = float(cue.get("No").strip())
        except (TypeError, AttributeError):
            sorszam = 0
        nev = cue.get("Name")
        note = cue.get("Note")
        cuefadein = ''
        cuefadeout = ''
        for child in cue:
            if "CueInFade" in child.attrib:
                cuefadein = child.attrib['CueInFade']
            if "CueOutFade" in child.attrib:
                cuefadeout = child.attrib['CueOutFade']
        trigtype = cue.get("TrigType")
        if trigtype is None:
            trigtype = "Go+"
        elif trigtype == "Time":
            ido = cue.get("TrigTime")
            trigtype = f"{trigtype} - {ido}"
        elif trigtype == "Sound":
            hang = cue.get("TrigSound")
            trigtype = f"{trigtype} - {hang}"
        comment = ""
        return [sorszam, nev, cuefadein, cuefadeout, note, trigtype, comment]

    for type_tag in root.iter("Cue"):
        treeData.append(process_cue(type_tag))

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = os.path.splitext(menu[vfile])[0]
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
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=4)
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
    xlsx_dir = "./xlsx"
    if not os.path.exists(xlsx_dir):
        print("xlsx directory not found, creating it.")
        os.makedirs(xlsx_dir)
    wb.save(f"{xlsx_dir}/{ws.title}.xlsx")
    wb.save(f"./xlsx/{ws.title}.xlsx")
    wb.close()
    print("Writing file done.\n Restarting")
