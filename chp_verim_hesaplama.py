import PySimpleGUI as sg
from openpyxl import Workbook, load_workbook

sg.theme("Dark")

excel_file = "CHP_verim.xlsx"
wb = load_workbook(excel_file)
ws = wb.active


layout = [

    [sg.Text("Please fill out information below")],
    
    [sg.Text("Ocak Ayı:", size=(15,1)), sg.InputText(key = "Ocak")],
    [sg.Text("Şubat Ayı:", size=(15,1)), sg.InputText(key = "Subat")],
    [sg.Text("Mart Ayı:", size=(15,1)), sg.InputText(key = "Mart")],
    [sg.Text("Nisan Ayı:", size=(15,1)), sg.InputText(key = "Nisan")],
    [sg.Text("Mayıs Ayı:", size=(15,1)), sg.InputText(key = "Mayis")],
    [sg.Text("Haziran Ayı:", size=(15,1)), sg.InputText(key = "Haziran")],
    [sg.Text("Temmuz Ayı:", size=(15,1)), sg.InputText(key = "Temmuz")],
    [sg.Text("Ağustos Ayı:", size=(15,1)), sg.InputText(key = "Agustos")],
    [sg.Text("Eylül Ayı:", size=(15,1)), sg.InputText(key = "Eylul")],
    [sg.Text("Ekim Ayı:", size=(15,1)), sg.InputText(key = "Ekim")],
    [sg.Text("Kasım Ayı:", size=(15,1)), sg.InputText(key = "Kasim")],
    [sg.Text("Aralık Ayı:", size=(15,1)), sg.InputText(key = "Aralik")],
    
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
        print(tuketim_list)
        
        i = 0
        for x in tuketim_list:
            a = ["E4","F4","G4","H4","I4","J4","K4","L4","M4","N4","O4","P4"]
            ws[a[i]].value = int(x)
            wb.save("CHP_verim.xlsx")
            i = i + 1
        sg.popup("Data Saved!")
        break


window.close()
