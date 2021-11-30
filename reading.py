import win32com.client
#pip install pywin32 - run this command for installing above package
import os

outlook = win32com.client.Dispatch("Outlook.Application").GetNameSpace("MAPI")

inbox = outlook.GetDefaultFolder(6)

message = inbox.Items
message2 = message.GetLast()
subject = message2.Subject
body = message2.body
date = message2.senton.date()
sender = message2.Sender
attachments = message2.Attachments
subject1 = 'A new requisition has been posted by Ingka Group SE this is for default'
subject2 = 'Closure Mails..'

for m in message:
    # if subject1 in m.Subject or subject2 in m.Subject:
    if subject1 in m.Subject:
        print(m.body)
