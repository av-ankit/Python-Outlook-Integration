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

i=0
for m in message:
    # if subject1 in m.Subject or subject2 in m.Subject:
    if subject1 in m.Subject:
        filename="raw_data_"+str(i)+".txt"
        text_file = open(filename, "w+")
        text_file.write(m.body)
        text_file.close()
        i+=1

for k in range(i):
    filename="raw_data_"+str(k)+".txt"
    file = open(filename, "r")
    contents = file.read()
    ind = contents.index("Requisition")
    contents=contents[ind:]
    contents = os.linesep.join([s for s in contents.splitlines() if s])

    file = open('data.csv', 'a', newline ='')   
    with file:
        write=csv.writer(file)
        for i in contents.split("\n"):
            i=i.strip()
            i=i.split(":")            
            write.writerow(i)
        write.writerow("\n")

    file.close()