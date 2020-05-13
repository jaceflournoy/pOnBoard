import PySimpleGUI as sg
import emails
import threading

layout = [sg.Button("Email 1"), sg.Button("Email 2"), sg.Button("Email 3")],\
         [sg.Button("Email 4"), sg.Button("Email 5"), sg.Button("Email 6")],\
         [sg.Multiline(key='emailStatus', size=(20, 5))],\
         [sg.Button("Close")]
window = sg.Window('Window Title', layout, keep_on_top=False, resizable=True)
while True:
    (event, value) = window.Read(timeout=100)
    if event ==  'EXIT'  or event is None or event == "Close":
        window.close()
        print("exit or none")
        break # exit button clicked
    elif event == "Email 1":
        window['emailStatus'].update("Opening Outlook for Email 1")
        e1t = threading.Thread(target=emails.email1, args=("email1.txt",))
        e1t.start()
        # emails.email1("email1.txt")
# To attach a file to the email (optional):
# attachment  = "Path to the attachment"
# mail.Attachments.Add(attachment)