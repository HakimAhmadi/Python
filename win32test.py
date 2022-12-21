import win32file
import win32inet
import win32profile
import win32ui
import win32wnet as wnet    
import win32api as beep
import win32api
import win32con
import win32com.client
import datetime

# file = win32file.FindFilesW('C:\\Users\\abdul\\Desktop\\*', Transaction=None)
# win32file.DecryptFile("C:\\Users\\abdul\\Desktop\\new.txt")
# print(file)
# for i in file:
#     print(i[-1])
# for i in range(10):
#     beep.Beep(500*i,500)
# print(win32profile.GetAllUsersProfileDirectory())
# print(win32ui.Is())
# string = wnet.WNetGetConnection()
# print(string)

# Reading cursor position
# while True:
#     pos = win32api.GetCursorPos()
#     print(pos)

# set cursor position
# pos = (200, 200)
# win32api.SetCursorPos(pos)


# def mouse_click(x, y):
#     win32api.SetCursorPos((x, y))
#     win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, x, y, 0, 0)
#     win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, x, y, 0, 0)

# while True:
#     mouse_click(20, 60)


  
# while 1:
#     print("Enter the word you want to speak it out by computer")
#     s = input()
#     speaker.Speak(s)
  





# the module to work with outlook

def speak(str):
    speaker = win32com.client.Dispatch("SAPI.SpVoice")
    speaker.Speak(str)

from enum import Enum

def reading_outlook_email():
        
    class OutlookFolder(Enum):

        outlookDeletedItemsFolder = 3 # The Deleted Items folder
        outlookOutboxFolder = 4 # The Outbox folder
        outlookSentMailFolder = 5 # The Sent Mail folder
        outlookInboxFolder = 6 # The Inbox folder
        outlookDraftsFolder = 16 # The Drafts folder
        outlookJunkFolder = 23 # The Junk E-Mail folder

    # get a reference to Outlook
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    # get the Inbox folder (you can a list of all of the possible settings at https://docs.microsoft.com/en-us/office/vba/api/outlook.oldefaultfolders)
    inbox = outlook.GetDefaultFolder(OutlookFolder.outlookInboxFolder.value)

    # get subfolder of this
    # todo = inbox.Folders.Item("To do - home")
    
    # get all the messages in this folder
    messages = inbox.Items

    # check messages exist
    if len(messages) == 0:

        speak("the searched folder is empty")
        exit()

    # loop over them all
    emails = []
    for message in messages:

    # get each email objects in a tuple    
        if message.Class == 43 and message.Senton.date() == datetime.date.today():
            this_message = (
                message.Subject,
                message.SenderEmailAddress,
                message.To,
                message.Unread,
                message.Senton.date(),
                message.body,
                message.Attachments
                )
        
            emails.append(this_message)
        
    #collecting all the emails that is of todays date on system
        # if this_message[4] == datetime.date.today():
        #     emails.append(this_message)
            
        # else:
        #     break
    

    speak(f"there are {len(emails)} emails for today" )
    
    for email in range(len(emails)):
        # text = f"email {email+1} is from {emails[email][1]} and subject is {emails[email][0]} and body message is {emails[email][5]}"
        text= f"email {email+1} is from {emails[email][1]} and subject is {emails[email][0]}"
        text+= f"Next"
        speak(text)
    exit()


    # To save the attachements fromt the email list
    # for email in emails:
    # subject, from_address, to_address, if_read, date_sent, body, attachments = email
    
    # # number of attachments
    #     if len(attachments) == 0:
    #         speak("No attachments")
    #     else:
    #         for attachment in attachments:
    #             attachment.SaveAsFile("c:\\Documents\\" + attachment.FileName)
    #             print("Saved {0} attachments".format(len(attachments)))
reading_outlook_email()