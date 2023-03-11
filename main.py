from pathlib import Path
import PySimpleGUI as sg
import pandas as pd
import csv

# ------- Helper Functions ------- #
def find_delimiter(filename):
    sniffer = csv.Sniffer()
    with open(filename) as fp:
        delimiter = sniffer.sniff(fp.read(5000)).delimiter
    return delimiter

def is_valid_path(filepath):
    if filepath and Path(filepath).exists():
        return True
    sg.popup("No file/folder selected!", title="Error", button_justification="center", button_color="red")
    return False

def display_file(file_path):
    try:
        if Path(file_path).suffix == ".xlsx":
            sheet_name = choose_sheet_name(pd.ExcelFile(file_path).sheet_names)
            df = pd.read_excel(file_path, sheet_name)
        else:
            df = pd.read_csv(file_path, index_col=False)
        filename = Path(file_path).name
        sg.popup_scrolled(df.dtypes, "=" * 50, df, title=filename)
    except:
        if sheet_name != "-CANCELLED-":
            sg.popup("Error parsing file", title="Error", button_justification="center", button_color="red")

def file_exists(file_folder, file_path):
    return Path(f'{file_folder}\{file_path}').is_file()

# ------- Replace File ------- #
def replace_file(file_path, file_type):
    result = True
    layout = [
        [sg.HSep()],
        [sg.Text(f"{Path(file_path)}{file_type} already exists. Replace?")],
        [sg.Button("Yes", key="-YES-", s=14, expand_x=True, button_color="tomato"), sg.Button("Cancel", key="-CANCEL-", s=14, expand_x=True)]
        ]
    window = sg.Window("Replace", layout, modal=True, use_custom_titlebar=True)
    while True:
        event, values = window.read()
        if event == sg.WINDOW_CLOSED or event == "-CANCEL-":
            result = False
            break
        if event == "-YES-":
            break
    window.close()
    return result

# ------- Choose Sheet ------- #
def choose_sheet_name(sheet_name):
    sheet_name_combo = '|'.join(sheet_name)
    layout = [
        [sg.Text("Sheet:"), sg.Combo(sheet_name_combo.split("|"), default_value=sheet_name[0], key="-SHEET-", readonly=True, expand_x=True)],
        [sg.Button("Yes", key="-YES-", s=14, expand_x=True, button_color="tomato"), sg.Button("Cancel", key="-CANCEL-", s=14, expand_x=True)],
        ]
    window = sg.Window("Choose a Sheet", layout, use_custom_titlebar=True, modal=True)
    while True:
        event, values = window.read()
        if event == sg.WINDOW_CLOSED or event == "-CANCEL-":
            sheet_name = "-CANCELLED-"
            break
        if event == "-YES-":
            sheet_name = values["-SHEET-"]
            break
    window.close()
    return sheet_name

# ------- Convert File ------- #
def convert_file(file_path, output_folder, separator, decimal, filetype, filename):
    result = True
    try:
        if filename == "":
            filename = Path(file_path).stem
        if filetype == 'xlsx':
            x1 = pd.ExcelFile(file_path)
            if file_exists(output_folder, filename + '.csv'):
                result = replace_file(filename, '.csv')
            if result == True:
                if len(x1.sheet_names) != 1:
                    sheet_name = choose_sheet_name(x1.sheet_names)
                if len(x1.sheet_names) == 1:
                    sheet_name = x1.sheet_names[0]
                if sheet_name == "-CANCELLED-":
                    int(sheet_name)
                df = pd.read_excel(file_path, sheet_name)
                outputfile = Path(output_folder) / f"{filename}.csv"
                df.to_csv(outputfile, sep=separator, decimal=decimal, index=False)
        if filetype == '.csv':
            if file_exists(output_folder, filename + '.xlsx'):
                result = replace_file(filename, '.xlsx')
            if result == True:
                delimiter = find_delimiter(file_path)
                df = pd.read_csv(file_path, sep=delimiter, index_col=0)
                outputfile = Path(output_folder) / f"{filename}.xlsx"
                df.to_excel(outputfile)
        if result: 
            sg.popup("Successfully Converted File!", button_justification="center", title="Success!")
    except:
        sg.popup("Conversion Cancelled.", button_justification="center", title="Cancelled")

# ------- Change Theme ------- #   
def change_theme(settings):

    # ------- GUI Definition ------- #
    layout = [
        [sg.HSep()],
        [sg.Text(f'Current theme: {GUI["theme"]}')],
        [sg.Combo(sg.theme_list(), s=33, key="-THEME-")],
        [sg.Button("OK", key="-OK-", s=16, button_color="tomato"), sg.Cancel(s=16)],
        [sg.HSep()],
        ]
    window = sg.Window("Theme", layout, modal=True, use_custom_titlebar=True)
    # ------- Event Loop ------- #
    while True:
        event, values = window.read()
        if event == sg.WINDOW_CLOSED or event == "Cancel":
            break
        if event == "-OK-":
            if len(values["-THEME-"]) != 0:
                GUI["theme"] = values["-THEME-"]
                sg.popup("Theme Changed! Restart required to take effect.", button_justification="center", title="Success!", modal=True)
                break
            else:
                sg.popup("Please make a selection.", button_justification="center", button_color="red", title="Error", modal=True)
    window.close()
        
# ------- Exit Program ------- #
def exit_program():
    
    # ------- GUI Definition ------- #
    layout = [
        [sg.HSep()],
        [sg.Text("Exit Program?")],
        [sg.Button("Yes", key="-YES-", s=16, button_color="tomato"), sg.Button("No", key="-NO-", s=16)],
        [sg.HSep()],
        ]
    window = sg.Window("Exit", layout, modal=True, use_custom_titlebar=True)
    # ------- Event Loop ------- #
    while True:
        event, values = window.read()
        if event == "-YES-":
            result = True
            break
        if event == "-NO-" or event == sg.WINDOW_CLOSED:
            result = False
            break
    window.close()
    return result

# ------- About Page ------- #
def about_page():
    layout = [
        [sg.HSep()],
        [sg.Text(GUI["title"], auto_size_text=True, justification="center", expand_x=True)],
        [sg.HSep()],
        [sg.Text("Source Code: github/chrisdmancuso", font=10, auto_size_text=True, justification="center", expand_x=True)],
        [sg.Text("Authored by Chris Mancuso, 2023", font=8, auto_size_text=True, justification="center", expand_x=True)],
        [sg.Text("Version 1.04", font=8, auto_size_text=True, justification="center", expand_x=True)],
        [sg.Text("Free to use", font=6, auto_size_text=True, justification="center", expand_x=True)],
        [sg.Button("OK", expand_x=True, key="-OK-")],
        [sg.HSep()],
        ]

    window = sg.Window("About", layout, use_custom_titlebar=True, modal=True)
    while True:
        event, values = window.read()
        if event == sg.WINDOW_CLOSED or event == "-OK-":
            break
    window.close()

# ------- Reset Defaults -------  #
def reset_defaults():
    layout = [
        [sg.HSep()],
        [sg.Text("Restore Theme/Settings to Default?", auto_size_text=True, expand_x=True, justification="center")],
        [sg.Button("Yes", key="-YES-", s=16, button_color="tomato"), sg.Button("No", key="-NO-", s=16)],
        [sg.HSep()],]
    window = sg.Window("Reset", layout, modal=True, use_custom_titlebar=True)
    # ------- Event Loop ------- #
    while True:
        event, values = window.read()
        if event == "-YES-":
            GUI["theme"] = "Reddit"
            CSV["separator"] = "|"
            CSV["decimal_default"] = "."
            CSV["sheet_name"] = "Sheet1"
            EXCEL["sheet_name"] = "Sheet1"
            sg.popup("Defaults Restored!", title="Success!", button_justification="center",)
            break
        if event == sg.WINDOW_CLOSED or event == "-NO-":
            break
    window.close()
    
# ------- Settings Window -------  #
def settings_window(settings):
    
    # ------- GUI Definition ------- #
    layout = [
        [sg.HSep()],
        [sg.Text("Decimal:", expand_x=True), sg.Combo(CSV["decimal"].split("|"), default_value=CSV["decimal_default"], s=1, key="-DECIMAL-", readonly=True)],
        [sg.Text("Delimiter:", expand_x=True), sg.Input(CSV["separator"], s=2, key="-DELIMITER-", justification="center", enable_events=True)],
        [sg.Text("Sheet Name:"), sg.Input(EXCEL["sheet_name"], s=20, key="-SHEET_NAME-", expand_x=True)],
        [sg.Button("Save Settings", key="-SAVE-", button_color="tomato", s=16, expand_x=True), sg.Button("Cancel", key="-CANCEL-", s=16, expand_x=True)],
        [sg.HSep()],
    ]
    window = sg.Window("Settings", layout, modal=True, use_custom_titlebar=True)
    # ------- Event Loop ------- #
    while True:
        event, values = window.read()
        if event == sg.WINDOW_CLOSED or event == "-CANCEL-":
            break
        if event == "-DELIMITER-" and len(values["-DELIMITER-"]) > 1:
            window["-DELIMITER-"].update(values["-DELIMITER-"][:-1])
        if event == "-SAVE-":
            if len(values["-DELIMITER-"]) == 1 and len(values["-SHEET_NAME-"]) != 0 and values["-DELIMITER-"].isalnum() == False:    
                CSV["separator"] = values["-DELIMITER-"]
                CSV["decimal_default"] = values["-DECIMAL-"]
                CSV["sheet_name"] = values["-SHEET_NAME-"]

                sg.popup("Settings saved!", button_justification="center", title="Success!", modal=True)
                break
            else:
                sg.popup("Delimiter/Sheet Name must have a value \n Delimiter cannot be alphanumeric (a-z, 0-9)", title="Error", button_justification="center", button_color="red", modal=True)
    window.close()

# ------- Main Window ------- #
def main_window():
    
    # ------- Menu Definition ------- #
    menu_def = [["File", ["Theme", "Settings", "Exit"]],
                ["Help", ["About", "Reset Defaults"]]]
    
    # ------- GUI Definition ------- #
    layout = [
        [sg.MenubarCustom(menu_def, tearoff=False)],
        [sg.HSep()],
        [sg.Text("Input File:", s=16, justification="r"), sg.Input(key="-IN-"), sg.FileBrowse(file_types=(("Excel Files", "*.xls*"), ("CSV", "*.csv*")))],
        [sg.Text("Output Folder:", s=16, justification="r"), sg.Input(key="-OUT-"), sg.FolderBrowse()],
        [sg.Text("(Optional) Filename:", s=16, justification="r"), sg.Input(key="-FILE-")],
        [sg.Exit(s=16, button_color="tomato", expand_x=True), sg.Button("Settings", s=16, expand_x=True), sg.Button("Display File", s=16, expand_x=True), sg.Button("Convert File", s=16, expand_x=True)],
    ]
    window_title = GUI["title"]
    #sg.LOOK_AND_FEEL_TABLE['Reddit']['ACCENT2']
    window = sg.Window(window_title, layout, enable_close_attempted_event=True)
    # ------- Event Loop ------- #
    while True:
        event, values = window.read()
        # --- Close or Exit --- #
        if event in (sg.WINDOW_CLOSE_ATTEMPTED_EVENT, "Exit"):
            result = exit_program()
            if result == True:
                break
        # --- About --- #
        if event == "About":
            about_page()
        # --- Theme --- #
        if event == "Theme":
            change_theme(settings)
        # --- Settings --- #
        if event == "Settings":
            settings_window(settings)
        # --- Reset --- #
        if event == "Reset Defaults":
            reset_defaults()
        # --- Display File --- #
        if event == "Display File":
            if is_valid_path(values["-IN-"]):
                display_file(
                    file_path=values["-IN-"])
        # --- Convert File --- #
        if event == "Convert File":
            if is_valid_path(values["-IN-"]) and is_valid_path(values["-OUT-"]):
                convert_file(
                    file_path=values["-IN-"],
                    output_folder=values["-OUT-"],
                    separator=CSV["separator"],
                    decimal=CSV["decimal_default"],
                    filetype=values["-IN-"][len(values["-IN-"]) - 4:],
                    filename=values["-FILE-"])
    window.close()

# ------- Main ------- #
if __name__ == "__main__":
    SETTINGS_PATH = Path.cwd()

    settings = sg.UserSettings(
        path=SETTINGS_PATH, filename="config.ini", use_config_file=True, convert_bools_and_none=True
        )
    GUI = settings["GUI"]
    CSV = settings["CSV"]
    EXCEL = settings["EXCEL"]
    
    theme = GUI["theme"]
    font_family = GUI["font_family"]
    font_size = int(GUI["font_size"])
    sg.theme(theme)
    sg.set_options(font=(font_family, font_size))
    main_window()

    
