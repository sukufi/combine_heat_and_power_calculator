import PySimpleGUI as sg
from openpyxl import Workbook, load_workbook

sg.theme("Dark")

# GİRİŞ İÇİN CELL VE EXCEL FİLE BELİRLE
cell_elc = ["E4","F4","G4","H4","I4","J4","K4","L4","M4","N4","O4","P4"]
cell_ngas = ["E5","F5","G5","H5","I5","J5","K5","L5","M5","N5","O5","P5"]
excel_file = "assets/template.xlsx"
output_excel_file = "CHP_verim.xlsx"

wb = load_workbook(excel_file)
ws = wb.active


layout = [

    [sg.Text("Please fill in the blanks with the \nmonthly electricity(kWh) and Natural Gas(m3) consumption")],
    [sg.Text("", size=(15,1)), sg.Text("Electricity", size=(15,1), justification='center'), sg.Text("Natural Gas", size=(15,1))],
    [sg.Text("January:", size=(15,1)), sg.InputText(key = "January_elc" , size=(15,1)), sg.InputText(key = "January_ngas", size=(15,1))],
    [sg.Text("February:", size=(15,1)), sg.InputText(key = "February_elc" , size=(15,1)), sg.InputText(key = "February_ngas", size=(15,1))],
    [sg.Text("March:", size=(15,1)), sg.InputText(key = "March_elc", size=(15,1)), sg.InputText(key = "March_ngas", size=(15,1))],
    [sg.Text("April:", size=(15,1)), sg.InputText(key = "April_elc", size=(15,1)), sg.InputText(key = "April_ngas", size=(15,1))],
    [sg.Text("May:", size=(15,1)), sg.InputText(key = "May_elc", size=(15,1)), sg.InputText(key = "May_ngas", size=(15,1))],
    [sg.Text("June:", size=(15,1)), sg.InputText(key = "June_elc", size=(15,1)), sg.InputText(key = "June_ngas", size=(15,1))],
    [sg.Text("July:", size=(15,1)), sg.InputText(key = "July_elc", size=(15,1)), sg.InputText(key = "July_ngas", size=(15,1))],
    [sg.Text("August:", size=(15,1)), sg.InputText(key = "August_elc", size=(15,1)), sg.InputText(key = "August_ngas", size=(15,1))],
    [sg.Text("September:", size=(15,1)), sg.InputText(key = "September_elc", size=(15,1)), sg.InputText(key = "September_ngas", size=(15,1))],
    [sg.Text("October:", size=(15,1)), sg.InputText(key = "October_elc", size=(15,1)), sg.InputText(key = "October_ngas", size=(15,1))],
    [sg.Text("November:", size=(15,1)), sg.InputText(key = "November_elc", size=(15,1)), sg.InputText(key = "November_ngas", size=(15,1))],
    [sg.Text("December:", size=(15,1)), sg.InputText(key = "December_elc", size=(15,1)), sg.InputText(key = "December_ngas", size=(15,1))],
    
    [sg.Text("Output:", size=(15,1)), sg.Input(key = "-FolderDirectory-", disabled= True, text_color="black", size=(32,1)), sg.FolderBrowse(button_text= "Browse", key= "Browse")],

    [sg.Submit(), sg.Button("Clear"), sg.Exit(),sg.Push(), sg.Button("Credits")]
    
        ]

window = sg.Window("Combined Heat And Power System Calculator", layout)

def clear_input():

    for key in value:
        if key != "Browse":
            window[key].update("")
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
    
    if event == "Credits":
        sg.popup("Desinged by \nKaan, Esat, Emre")

    if event == "Submit":
        
        output_folder = value["-FolderDirectory-"]
        value.pop("Browse")
        value.pop("-FolderDirectory-")
        
        if value_check(value) == [] and output_folder != "":
            
            save_spot = output_folder + "/" + output_excel_file
            all_usage = list(value.values())
            i = 0
            j = 0
            for x in all_usage:
                if i % 2 == 0:
                    ws[cell_elc[j]].value = int(x)
                    wb.save(save_spot)
                else:
                    ws[cell_ngas[j]].value = int(x)
                    wb.save(save_spot)
                    j = j + 1
                i = i + 1

            
            sg.popup("Data Saved!")
            window.close()

        else:
            
            
            wrong_entry = value_check(value)
            
            if output_folder == "":
                wrong_entry.append("Output Directory")

            sg.popup(f"The following sections have been filled in incorrectly or incompletely, please try again.\n\n--> {', '.join(wrong_entry)}")
               
               
            
