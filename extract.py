import imaplib
import email
import os
from datetime import date
# onDate = "SENTON 05-Mar-2007"
today = date.today()
d1 = today.strftime("%d-%B-%Y")
print(d1)
server = 'outlook.office365.com' #Enter the email server
user = ''     #USERNAME
password = '' #PASSWORD
outputdir = ''   #OUTPUT DIRECTORY
subject = 'SPNA INVENTORY HISTORY' #subject line of the emails you want to download attachments from

# connects to email client through IMAP
def connect(server, user, password):
    m = imaplib.IMAP4_SSL(server)
    m.login(user, password)
    m.select()
    return m



def downloaAttachmentsInEmail(m, emailid, outputdir):
    resp, data = m.fetch(emailid, "(BODY.PEEK[])")
    print(emailid)
    email_body = data[0][1]
    mail = email.message_from_bytes(email_body)
    if mail.get_content_maintype() != 'multipart':
        return
    for part in mail.walk():
        if part.get_content_maintype() != 'multipart' and part.get('Content-Disposition') is not None:
            open(outputdir + '/' + d1+".zip", 'wb').write(part.get_payload(decode=True))


def subjectQuery(subject):
    m = connect(server, user, password)
    m.select("Inbox")
    typ, msgs = m.search(None, '(SUBJECT "' + subject +'" SENTON "'+d1+'")')
    msgs = msgs[0].split()
    print(msgs)
    for emailid in msgs:
        downloaAttachmentsInEmail(m, emailid, outputdir)

subjectQuery(subject)


from zipfile import ZipFile

zf = ZipFile(d1+".zip", 'r')
print(zf.namelist())
newfile = zf.open(zf.namelist()[0])
zf.extractall('/safal')


import csv

data = []

with open(zf.namelist()[0], 'r') as f:
    reader = csv.reader(f, dialect = 'excel', delimiter = '\t')
    # % reads the rows from your imported data file and appends them to a list
    for row in reader:
        # print(row)
        data.append(row)

zf.close()
import csv
import pandas as pd
with open(d1+'.csv', 'a+', newline='') as file:
    fieldnames = ['c1','c2','c3','c4','c5']
    writer = csv.DictWriter(file, fieldnames=fieldnames)
    for i in range(0,len(data)):
        writer.writerow({'c1':data[i][0],'c2':data[i][1],'c3':data[i][2],'c4':data[i][3],'c5':data[i][4]})
