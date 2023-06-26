import win32com.client
#other libraries to be used in this script
import os
from datetime import datetime, timedelta, date
import pandas as pd
from dateutil.parser import *
import pandas as pd

################################################################
# Filter emails Messages Function
################################################################
def filterMessages(messages):
    received_dt = datetime.now() - timedelta(days=5)
    received_dt = received_dt.strftime('%m/%d/%Y %H:%M %p')
    #messages = messages.Restrict("[SenderEmailAddress] = 'teste@email.com'")
    messages = messages.Restrict("[ReceivedTime] >= '" + received_dt + "'")
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
    inbox = mapi.GetDefaultFolder(9)
    appts = inbox.Items
    return appts

def generateReportFromMessages(messages):
    d = []
    for message in list(messages):
        d.append(
            {
                #'fields': message.__dir__(),
                'sender': message.SenderName,
                'senderemailaddress': message.SenderEmailAddress,
                'read': message.UnRead,
                'subject': message.Subject,
                'receivedtime': message.ReceivedTime
            }
        )
    df = pd.DataFrame(d)
    df.to_csv('report_emails.csv', sep=";", encoding='utf-8-sig')  

if __name__ == "__main__":
    print("Start main function 1")
    appts = getOutlookClient()
    print(len(appts))
    ##for app in appts:
##        print(app.Start)
    excluded_subjects=('Cancelado')
    # Step 1, block 2 : sort events by occurrence and include recurring events
    appts.Sort("[Start]")
    appts.IncludeRecurrences = "True"
    #2023-05-22 11:30:00+00:00
    end = date.today() + timedelta(days=2)
    end = end.strftime("%m/%d/%Y")
    begin = date.today() - timedelta(days=1)
    #begin = date.today() - timedelta(days=5)
    begin = begin.strftime("%m/%d/%Y")

    print(type(begin))

    appts = appts.Restrict("[Start] >= '" +begin+ "' AND [END] <= '" +end+ "'")
    #appts = appts.Restrict("[Start] >= '" +begin+ "'")
    apptDict = {}
    d = []
    item = 0
    for indx, a in enumerate(appts):
        subject = str(a.Subject)
#        print(subject)
#        if subject in (excluded_subjects) or "Cancelado" in subject:
        if "Cancelado" in subject or "Canceled" in subject:
            continue
        else:
            meetingDate = str(a.Start.replace(tzinfo=None))
            d.append(
                {
                'fields': a.__dir__(),
                'Organizer': a.Organizer,
                'Start': a.Start.replace(tzinfo=None),
                'Subject': a.Subject,
                'Duration': a.duration,
                'Date':parse(meetingDate).date(),
                'IsOnlineMeeting': a.IsOnlineMeeting
                #'Conflicts': a.Conflicts.Count
                }
                )
            #organizer = str(a.Organizer)
            #meetingDate = str(a.Start)
            #date = parse(meetingDate).date()
            #subject = str(a.Subject) 
            #duration = str(a.duration)
            #apptDict[item] = {"Duration": duration, "Organizer": organizer, "Subject": subject, "Date": date.strftime("%m/%d/%Y")}
            #item = item + 1
    print("End main function")
    df = pd.DataFrame(d)
    #df.to_csv('report_meeting.csv', sep=";")  
    df.to_csv('report_meeting.csv', sep=";", encoding='utf-8-sig')

    df = df[['Duration', 'Organizer', 'Subject', 'Date']]
    apt_df = df.set_index('Date')
    apt_df['Duration'] = apt_df['Duration'].astype(str)
    apt_df['Meetings'] = apt_df[['Duration', 'Organizer', 'Subject']].agg(' | '.join, axis=1)
    grouped_apt_df = apt_df.groupby('Date').agg({'Meetings':', '.join})
    grouped_apt_df.index = pd.to_datetime(grouped_apt_df.index)
    grouped_apt_df.sort_index()
    print(grouped_apt_df)
    grouped_apt_df.to_csv('report_grouped_meeting.csv', sep=";", encoding='utf-8-sig')

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
