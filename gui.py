import PySimpleGUI as sg
import time
from Convertor import convert

clock_label = sg.Text(key="clock")
input_label = sg.Text("Enter the Excel file: ")
intput_box = sg.Input(key="File input box")
input_button = sg.FilesBrowse(key="Select files")
message_label = sg.Text(key="message")
convert_button = sg.Button("Convert")
exit_button = sg.Button("Exit")
layout = [[sg.VPush()],
          [clock_label],
          [input_label, intput_box, input_button],
          [message_label],
          [sg.Push(), convert_button, exit_button, sg.Push()],
          [sg.VPush()]]

# Initialize the window
window = sg.Window("Excel to PDF Convertor", layout=layout, font=("Helvetica", 20))

while True:
    event, values = window.read(timeout=200)

    window["clock"].update(value=time.strftime("%b %d, %Y %H:%M:%S"))
    match event:
        case "Convert":
            for filepath in values["Select files"].split(";"):
                convert(filepath)
            window["message"].update("Files converted successfully!")
            window["File input box"].update("")

        case "Exit":
            break

        case sg.WINDOW_CLOSED:
            break

window.close()






