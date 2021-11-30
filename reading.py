import win32com.client,os,csv
#pip install pywin32 - run this command for installing above package

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
        text_file = open("raw_data.txt", "a")
        text_file.write(m.body)
        text_file.close()

file = open("raw_data.txt", "r")

contents = file.read()
ind = contents.index("Requisition")
contents=contents[ind:]
contents = os.linesep.join([s for s in contents.splitlines() if s])
# print(contents)

file = open('data.csv', 'w+', newline ='')   
with file:
    for i in contents.split("\n"):
        i=i.strip()
        i=i.split(":")
    
        write=csv.writer(file)
        write.writerow(i)

file.close()



