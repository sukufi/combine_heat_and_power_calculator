import re
import PySimpleGUI as sg
from openpyxl import Workbook, load_workbook

sg.theme("Dark")

excel_file = "CHP_verim.xlsx"
wb = load_workbook(excel_file)
ws = wb.active


layout = [

    [sg.Text("Please fill in the blanks with the monthly electricity consumption in kWh.")],
    
    [sg.Text("January:", size=(15,1)), sg.InputText(key = "Ocak")],
    [sg.Text("February:", size=(15,1)), sg.InputText(key = "Subat")],
    [sg.Text("March:", size=(15,1)), sg.InputText(key = "Mart")],
    [sg.Text("April:", size=(15,1)), sg.InputText(key = "Nisan")],
    [sg.Text("May:", size=(15,1)), sg.InputText(key = "Mayis")],
    [sg.Text("June:", size=(15,1)), sg.InputText(key = "Haziran")],
    [sg.Text("July:", size=(15,1)), sg.InputText(key = "Temmuz")],
    [sg.Text("August:", size=(15,1)), sg.InputText(key = "Agustos")],
    [sg.Text("September:", size=(15,1)), sg.InputText(key = "Eylul")],
    [sg.Text("October:", size=(15,1)), sg.InputText(key = "Ekim")],
    [sg.Text("November:", size=(15,1)), sg.InputText(key = "Kasim")],
    [sg.Text("December:", size=(15,1)), sg.InputText(key = "Aralik")],
    
    [sg.Submit(), sg.Button("Clear"), sg.Exit()]
]

window = sg.Window("Combined Heat And Power System Calculator", layout)

def clear_input():
    for key in value:
        window[key]("")
    return None


while True:
    event, value = window.read()
    
    if event == sg.WIN_CLOSED or event == "Exit":
        break
    
    if event == "Clear":
        clear_input()

    if event == "Submit":
        
        tuketim_list = list(value.values())
        
        i = 0
        for x in tuketim_list:
            
            x = x.lstrip().rstrip()
            
            if re.match("^\d*$", x) == None or x == '' :
                print("failed")
                sg.popup("You must:\n-Fill all blanks\n-Use digits only")
                quit()
            else:
                print("passed")
                a = ["E4","F4","G4","H4","I4","J4","K4","L4","M4","N4","O4","P4"]
                ws[a[i]].value = int(x)
                print(x)
                wb.save("CHP_verim.xlsx")
                i = i + 1

        sg.popup("Data Saved!")
        break


window.close()
