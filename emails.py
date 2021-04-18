import win32com.client as win32
import pythoncom

# outlook = win32.Dispatch('outlook.application')

def email1(filename):
    pythoncom.CoInitialize()
    outlook = win32.Dispatch('outlook.application')
    file = open("./Templates/" + filename, "r")
    contents = file.read()
    mail = outlook.CreateItem(0)
    mail.To = 'test@test.com'
    mail.CC = "test@test.com"
    mail.Subject = 'Message subject'
    mail.Body = contents
    mail.Display(True) # open outlook before sending
    # mail.Send() # send email automatically through python

