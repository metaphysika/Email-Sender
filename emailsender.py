import win32com.client as win32
import psutil
import os

'''
This class is meant for import as a module to send emails.  Attachments are optional arguemnts.
Need to add a file path if attachemnts are wanted.

Multiple email addresses can be used by separating addresses with ; inside one set of quotes.
ex. "email.1@business.org; email.2@business.edu"
'''


class EmailSender():

    def __init__(self):
        pass

    def send_email(self, to, subject, body, attachment=None):
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = to
        mail.Subject = subject
        mail.body = body
        try:
            mail.Attachments.Add(attachment)
        except:
            pass
        mail.send

    # Checking if outlook is already opened. If not, open Outlook.exe

    def check_outlook(self):
        for item in psutil.pids():
            p = psutil.Process(item)
            flag = (p.name() == "OUTLOOK.EXE")
            if flag:
                break

        if flag:
            pass
        else:
            try:
                os.startfile("outlook")
            except:
                print("Outlook didn't open successfully")
