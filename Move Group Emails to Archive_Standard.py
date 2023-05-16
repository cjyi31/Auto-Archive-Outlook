# -*- coding: utf-8 -*-
"""
Created on Thu Nov 24 14:36:34 2022

@author: Grace.Choo
"""

"""
This is to archive emails dated > 90 days.
"""

####################
#    0. Settings   #
####################
#to move emails > x number days old to archive folders and subfolders
delta = 90

##########################
#    0. Import Macros    #
##########################
exec(open(r'C:\Users\Documents\Common Macros.py').read())

#Start the timer
Macro_CountTimerStarts()

############################
#    0a. Import packages   #
############################
import win32com.client as client

import pandas as pd

import datetime as dt
from datetime import date
from datetime import datetime
from dateutil.relativedelta import relativedelta

#to move emails > 90 days old to archive
MaxCutOffDate = date.today() + relativedelta(days=-delta)
MaxCutOffDate = datetime.combine(MaxCutOffDate, datetime.max.time())


currentDT = dt.datetime.now()
#output file name
filename=r"C:\Users\Documents\Archive Log\ArchiveLog_"+ str(currentDT.strftime("%Y%m%d_%H%M%S")) +".xlsx"

#############################
#    1. Initiate outlook    #
#############################

outlook=client.Dispatch("Outlook.Application")
nameSpace=outlook.GetNameSpace("MAPI")

#Change your mailbox name here:
MailBox = nameSpace.folders['Mailbox Name']

#To check location of Inbox folder
i=0
Inbox_Num = -1 #initialized at -1
Archive_Num = -1 #initialized at -1

print("list of outlook email account:")
for folder in MailBox.Folders:
    #print(str(i) + ": " + folder.Name)
    #assign folder number it may change. As such, need to create a way to automate
    #Note:
        #Archive = 21
        #Inbox = 1
    if folder.Name == "Inbox":
        Inbox_Num = i
    
    if folder.Name == "Archive":
        Archive_Num = i
    i+=1

#Assign Inbox and Archive
MailBox_Inbox = MailBox.Folders[Inbox_Num]
MailBox_Archive = MailBox.Folders[Archive_Num]

#print(MailBox_Inbox.Name)
#print(MailBox_Archive.Name)


print("Num of Emails in " + MailBox_Inbox.Name +" folder is: "+ str(MailBox_Inbox.Items.Count))
print("Num of Emails in " + MailBox_Archive.Name +" folder is: "+ str(MailBox_Archive.Items.Count))



######################################################################
#    2. inbox subfolder move to Archive Subfolder of the same name   #
######################################################################
def MoveEmailsToCorrectSubfolders(i,j):
    global RecallNum # do not use this lightly, only use when it is absoltely nessasary(for UnboundLocalError)
    k = 0
    j = 0 # this is to track how manytimes k is reset whenever the email sent date > MaxCustOffDate
    while MailBox_Inbox.Folders[i].Items.Count > 0:
        a = MailBox_Inbox.Folders[i].Items.Count - 1 - k
        try:
            print("Looking at Email: " + MailBox_Inbox.Folders[i].Items[a].Subject)
            print(MailBox_Inbox.Folders[i].Items[a].SentOn.strftime("%Y-%m-%d %H:%M:%S"))
            if datetime.strptime(MailBox_Inbox.Folders[i].Items[a].SentOn.strftime("%Y-%m-%d %H:%M:%S"), '%Y-%m-%d %H:%M:%S') <= MaxCutOffDate:
                print("Email Sent datetime is:")
                print(datetime.strptime(MailBox_Inbox.Folders[i].Items[a].SentOn.strftime("%Y-%m-%d %H:%M:%S"), '%Y-%m-%d %H:%M:%S'))
                print("")
                print("Moving: " + MailBox_Inbox.Folders[i].Items[a].Subject)
                print("")

                #keep a log
                RowbyRow = pd.DataFrame({'Subfolder' : MailBox_Inbox.Folders[i].Name, 'EmailTitle' : MailBox_Inbox.Folders[i].Items[a].Subject, 'EmailSentDateTime' : datetime.strptime(MailBox_Inbox.Folders[i].Items[a].SentOn.strftime("%Y-%m-%d %H:%M:%S"), '%Y-%m-%d %H:%M:%S'),
                                         'EmailSender' : MailBox_Inbox.Folders[i].Items[a].SenderName, """'EmailSenderAdd' : EmailSenderAdd,""" 'Session' : currentDT.strftime("%Y%m%d %H:%M:%S")}, index=[0])
                #output and keep each results
                Record.append(RowbyRow)
                
                print(MailBox_Inbox.Folders[i].Items[a].Subject + " moved to archive folder.")
                MailBox_Inbox.Folders[i].Items[a].Move(MailBox_Archive.Folders[j])

            #if date is out of range then proceed to the next inbox subfolder
            elif (datetime.strptime(MailBox_Inbox.Folders[i].Items[a].SentOn.strftime("%Y-%m-%d %H:%M:%S"), '%Y-%m-%d %H:%M:%S') > MaxCutOffDate ) & (j <= 5):
                j += 1
                print("")
                print("====== j = " + str(j) +" times. Reset index k to 0. Continue searching within the same subfolder. ======")
                print("")
                print("====== Current Subfolder is " + MailBox_Inbox.Folders[i].Name + "======")
                print("")
                k = 0 #reset to 0 because once emails are moved, the index of each email will change.
                continue
            
            #if date is out of range then proceed to the next inbox subfolder
            elif (datetime.strptime(MailBox_Inbox.Folders[i].Items[a].SentOn.strftime("%Y-%m-%d %H:%M:%S"), '%Y-%m-%d %H:%M:%S') > MaxCutOffDate ) & (j > 5):
                print("")
                print("====== Go to the next subfolder ======")
                print("")
                j = 0 #reset j = 0
                k = 0
                break
            
            #set up a kill command. because if not, the MoveEmailsToCorrectSubfolders macro will run forever (set the stop comman at minimum of remaining number of emails to 100)
            if k == 999:
                print("Hit the kill command")
                break
                
        except IndexError:
            print("Index Error. Reset index to 0")
            k = 0 #reset to 0 because once emails are moved, the index of each email will change.
            continue
        
        except:
            if RecallText in MailBox_Inbox.Folders[i].Items[a].Subject:
                print("have recall error. Reset index to 1")
                RecallNum += 1 #reset to 1 because when there is an email recalled, it will stuck in a loop
                k = RecallNum
                continue
        k += 1



###
#To look for same subfolder name
####
    
#Set recall text:
RecallText = "Recall: "
Record = []
a = -1
b = -1
k = 0
RecallNum = 0

Inbox_Num2 = -1
Archive_Num2 = -1
for a in range(MailBox_Inbox.Folders.Count):
    for b in range(MailBox_Archive.Folders.Count):
        if MailBox_Archive.Folders[b].Name == MailBox_Inbox.Folders[a].Name:
            Inbox_Num2 = a
            print("Current Inbox Subfolder is: " + MailBox_Inbox.Folders[a].Name)
            Archive_Num2 = b
            print("Current Archive Subfolder is: " + MailBox_Archive.Folders[b].Name)
            #reset RecallNum = 0
            RecallNum = 0
            MoveEmailsToCorrectSubfolders(Inbox_Num2, Archive_Num2)





print("")     
print("====== Move Emails Completed! Here's a Christmass Tree for you to celebrate:")
print("")
print("")


ChristmasTreePattern(10)

print("")
print("")




# Write the full results to csv using the pandas library.
ArchiveLog = pd.concat(Record)

#output log as excel file
ArchiveLog.to_excel(filename, index=False, header=True)


#Time end and find out how long the process takes
Macro_CountTimerEnds()










