import PySimpleGUI as sg
from openpyxl import Workbook, load_workbook

sg.theme("Dark")

# GİRİŞ İÇİN CELL VE EXCEL FİLE BELİRLE
cell = ["E4","F4","G4","H4","I4","J4","K4","L4","M4","N4","O4","P4"]
excel_file = "CHP_verim.xlsx"
output_excel_file = "CHP_verim.xlsx"

wb = load_workbook(excel_file)
ws = wb.active


layout = [

    [sg.Text("Please fill in the blanks with the monthly electricity consumption in kWh.")],
    
    [sg.Text("January:", size=(15,1)), sg.InputText(key = "January")],
    [sg.Text("February:", size=(15,1)), sg.InputText(key = "February")],
    [sg.Text("March:", size=(15,1)), sg.InputText(key = "March")],
    [sg.Text("April:", size=(15,1)), sg.InputText(key = "April")],
    [sg.Text("May:", size=(15,1)), sg.InputText(key = "May")],
    [sg.Text("June:", size=(15,1)), sg.InputText(key = "June")],
    [sg.Text("July:", size=(15,1)), sg.InputText(key = "July")],
    [sg.Text("August:", size=(15,1)), sg.InputText(key = "August")],
    [sg.Text("September:", size=(15,1)), sg.InputText(key = "September")],
    [sg.Text("October:", size=(15,1)), sg.InputText(key = "October")],
    [sg.Text("November:", size=(15,1)), sg.InputText(key = "November")],
    [sg.Text("December:", size=(15,1)), sg.InputText(key = "December")],
    
    [sg.Text("Output:"), sg.Input(key = "-FolderDirectory-"), sg.FolderBrowse()],

    [sg.Submit(), sg.Button("Clear"), sg.Exit()]
    
        ]

window = sg.Window("Combined Heat And Power System Calculator", layout)

def clear_input():
    for key in value:
        window[key]("")
    return None

def value_check(dict):
    
    key = list(dict.keys())
    usage = list(dict.values())
    wrong_months = []
    
    i = 0
    for x in usage:
        try:
            y = float(x)
            print(f"{y} passed", end=", ")
            i = i + 1
        except ValueError:
            wrong_months.append(key[i])
            print(f"{x} failed", end=", ")
            i = i + 1
    print("\n------------------------------")
    return wrong_months


while True:
    event, value = window.read()

    if event == sg.WIN_CLOSED or event == "Exit":
        break
    
    if event == "Clear":
        clear_input()

# KULLANICI BOŞ VEYA HATALI GİRİŞ YAPACAK HANGİ AYLARIN YANIŞ GİRİLDİĞİNİ SÖYLEYECEK TEKRAR GİRİŞ YAPMASINI İSTEYECE

    if event == "Submit":
        
        output_folder = value["-FolderDirectory-"]
        value.pop("Browse")
        value.pop("-FolderDirectory-")
        
        if value_check(value) == [] and output_folder != "":
            
            save_spot = output_folder + "/" + output_excel_file
            electric_usage = list(value.values())
            i = 0
            for x in electric_usage:
                ws[cell[i]].value = int(x)
                wb.save(save_spot)
                i = i + 1
            
            sg.popup("Data Saved!")
            window.close()

        else:
            
            wrong_entry = value_check(value)
            
            if output_folder == "":
                wrong_entry.append("Output Directory")

            sg.popup(f"The following sections have been filled in incorrectly or incompletely, please try again.\n\n--> {', '.join(wrong_entry)}")
               
               
            
