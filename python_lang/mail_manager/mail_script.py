import win32com.client
#other libraries to be used in this script
import os
from datetime import datetime, timedelta
import pandas as pd


def sendMeeting():    
  outlook = win32com.client.Dispatch('outlook.application')
  appt = outlook.CreateItem(1) # AppointmentItem
  appt.Start = "2023-05-23 15:10" # yyyy-MM-dd hh:mm
  appt.Subject = "Subject of the meeting"
  appt.Duration = 60 # In minutes (60 Minutes)
  appt.Location = "Location Name"
  appt.MeetingStatus = 1 # 1 - olMeeting; Changing the appointment to meeting. Only after changing the meeting status recipients can be added
  appt.IsOnlineMeeting = True
  
  #appt.Recipients.Add("test@test.com") # Don't end ; as delimiter

  appt.Save()
  appt.Send()

def sendEmail(subject, mailto, body, attachment_path):
    outlook = win32com.client.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = mailto
    mail.Subject = subject
    mail.Body = body
    #mail.HTMLBody = '<h2>HTML Message body</h2>' #this field is optional
    # To attach a file to the email (optional):
    attachment  = attachment_path
    mail.Attachments.Add(attachment)
    mail.Send()

################################################################
# Filter emails Messages Function
################################################################
def filterMessages(messages,delta_days=5):
    received_dt = datetime.now() - timedelta(days=delta_days)
    received_dt = received_dt.strftime('%m/%d/%Y %H:%M %p')
    #messages = messages.Restrict("[SenderEmailAddress] = 'teste@email.com'")
    messages = messages.Restrict("[ReceivedTime] >= '" + received_dt + "'")
    messages = messages.Restrict("[UnRead] = True")
    return messages

################################################################
# Deletando emails Function
################################################################

def deleteEmail(messages):
    for message in list(messages):
        print("Deleting email {}".format(message))
        message.Delete()
        #print(message.SenderEmailAddress)
        
################################################################
# Print emails Function
################################################################
def listFolders(directoryFolders):
    for folder in directoryFolders.Folders: 
        print(folder.Name)

################################################################
# Get Client from Outlook
################################################################
def getOutlookClient():
    outlook = win32com.client.Dispatch('outlook.application')
    mapi = outlook.GetNamespace("MAPI")
    inbox = mapi.GetDefaultFolder(6)
    messages = inbox.Items
    return messages

def generateReportFromMessages(delta_days=5):
    messages = getOutlookClient()
    messages = filterMessages(messages,delta_days)
    d = []
    for message in list(messages):
        if message.Class == 43:
            d.append(
                {
                    #'fields': message.__dir__(),
                    'sender': message.SenderName,
                    'senderemailaddress': message.SenderEmailAddress,
                    'read': message.UnRead,
                    'subject': message.Subject,
                    'receivedtime': message.ReceivedTime.replace(tzinfo=None)
                }
            )
    df = pd.DataFrame(d)
    df.to_csv('report_emails.csv', sep=";", encoding='utf-8-sig')  

if __name__ == "__main__":
    #print("Start main function")
    #messages = getOutlookClient()
    #messages = filterMessages(messages)
    #generateReportFromMessages(messages)
    #sendEmail()
    print("End main function")


#########################################################################

# SandBox

#########################################################################

        #if message.SenderName == "[Cortex]":
            #d.append(
                #{
                    #'sender': message.SenderName,
                    #'read': message.UnRead,
                    #'subject': message.Subject,
                    #'ReceivedTime': message.ReceivedTime
                #}
            #)
    #        message.Delete()
#messages = messages.Restrict("[Subject] = 'Assunto Email'")
#messages = messages.Restrict("[HasAttachment] = 'yes'")

#for folder in inbox.Folders: 
#    print(folder.Name)
#directoryFolders = inbox.Folders['Projetos'].Folders['Pasta']

#directoryFolders.Folders.Add("`Pasta")
#directoryFolders = inbox.Folders['Projetos']

#listFolders(directoryFolders)
#for message in list(messages):
#    messages.move(directoryFolders)
    #print(message.Body)
    
#messages.move(directoryFolders)
#for message in list(messages):
#    message.move()
