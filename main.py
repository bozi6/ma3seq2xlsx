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


def get_default_sequence_directory():
    """Get the default MA3 sequence directory based on the operating system."""
    if platform.system() == "Windows":
        return "C:\\ProgramData\\MALightingTechnology\\gma3_library\\datapools\\sequences\\"
    elif platform.system() == "Darwin":
        homedir = os.path.expanduser("~")
        return f'{homedir}/MALightingTechnology/gma3_library/datapools/sequences/'
    else:
        return ""


def initialize_application():
    """Initialize the application and return menu dictionary."""
    default_dir = get_default_sequence_directory()

    if not os.path.isdir(default_dir):
        print(f"Error: Directory '{default_dir}' does not exist or is inaccessible.")
        exit(1)

    filenames = os.listdir(default_dir)
    menu = {}
    menunum = 0

    for filename in filenames:
        if filename.endswith(".xml"):
            menunum += 1
            menu[menunum] = filename

    menu[len(menu) + 1] = "Exit"
    return menu, default_dir


def menushow(menu):
    """Display a formatted menu to the console."""
    print("-" * 30)
    print("MENU")
    print("-" * 30)
    for key in menu.keys():
        print(key, "--", menu[key])


def process_cue(cue):
    """Process a single cue element and return its data."""
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


def process_xml_file(file_path, filename):
    """Process XML file and return sequence data."""
    print("XML file: ", file_path)
    root = et.parse(file_path).getroot()
    verzio = root.attrib["DataVersion"]

    seq_element = root.find(".//Sequence")
    if seq_element is not None and "Note" in seq_element.attrib:
        seqnote = seq_element.attrib["Note"].replace("&#xD;", " ")
        seqnote = seqnote.replace("\r", " ")
        seqnote = seqnote.strip()
    else:
        seqnote = ""

    tree_data = [
        ["Num:", "Cue name:", "FadeIn:", "FadeOut", "Cue Note:", "Trig.Type/Param:", "Comment:"]
    ]

    for cue in root.iter("Cue"):
        tree_data.append(process_cue(cue))

    return tree_data, verzio, seqnote, os.path.splitext(filename)[0]


def create_excel_file(tree_data, verzio, seqnote, title):
    """Create and save Excel file with sequence data."""
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = title

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

    for tree in tree_data:
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
    wb.close()


def run_application():
    """Main application logic."""
    menu, default_dir = initialize_application()

    running = True
    while running:
        menushow(menu)
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

        file_path = default_dir + menu[vfile]
        tree_data, verzio, seqnote, title = process_xml_file(file_path, menu[vfile])
        create_excel_file(tree_data, verzio, seqnote, title)
        print("Writing file done.\n Restarting")


if __name__ == "__main__":
    run_application()
