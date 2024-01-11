import PySimpleGUI as sg
import pandas as pd

sg.theme("Default1")

layout = [
    [sg.Text("Select Excel file:")],
    [sg.Input(key="-FILEPATH-", enable_events=True), sg.FileBrowse()],
    [sg.Text("Select sheets to convert:")],
    [sg.Listbox(values=[], size=(30, 6), key="-SHEETS-", enable_events=True, select_mode=sg.LISTBOX_SELECT_MODE_MULTIPLE)],
    [sg.Button("Update sheets", key="-UPDATE_SHEETS-")],
    [sg.Text("Select output folder for PDF files:")],
    [sg.Input(key="-OUTPUT_DIR-", enable_events=True), sg.FolderBrowse()],
    [sg.Text("Enter cell for PDF file name (optional):")],
    [sg.Input(key="-CELL-", size=(10, 1))],
    [sg.Button("Convert to PDF", key="-CONVERT-")]
]

window = sg.Window("Excel to PDF Converter", layout, finalize=True)

while True:
    event, values = window.read()

    if event == sg.WIN_CLOSED:
        break

    if event == "-FILEPATH-":
        file_path = values["-FILEPATH-"]
        if file_path:
            df = pd.ExcelFile(file_path)
            sheet_names = df.sheet_names
            window["-SHEETS-"].update(sheet_names)


    if event == "-UPDATE_SHEETS-":
        file_path = values["-FILEPATH-"]
        if file_path:
            df = pd.ExcelFile(file_path)
            sheet_names = df.sheet_names
            window["-SHEETS-"].update(sheet_names)

window.close()