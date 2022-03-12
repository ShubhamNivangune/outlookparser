import win32com.client
import os
import shutil
from tkinter import * 
from tkinter import messagebox
from datetime import datetime, timedelta
#----------------------------------Import lib-----------------------------------

if os.path.exists(os.getcwd() + "/attachments") :
    outlook = win32com.client.Dispatch('outlook.application')
    mapi = outlook.GetNamespace("MAPI")
    for account in mapi.Accounts:
        print(account.DeliveryStore.DisplayName)
    inbox = mapi.GetDefaultFolder(6)
    messages = inbox.Items
    #Let's assume we want to save the email attachment to the below directory
    outputDir = r""+os.getcwd()+"\\attachments"
    try:
        for message in list(messages):
            try:
                s = message.sender
                for attachment in message.Attachments:
                    attachment.SaveASFile(os.path.join(outputDir, attachment.FileName))
                    print(f"attachment {attachment.FileName} from {s} saved")
                    
            except Exception as e:
                print("error when saving the attachment:" + str(e))
    except Exception as e:
            print("error when processing emails messages:" + str(e))
    if os.path.exists(os.getcwd() + "/resumes") :
        arr = os.listdir(r""+os.getcwd()+"\\attachments")
        for i in range(len(arr)):
            if(('.docx') in arr[i] or ('.doc') in arr[i] or ('.pdf') in arr[i]  ):
                original = r""+os.getcwd()+"\\attachments\\"+arr[i]
                target = r""+os.getcwd()+"\\resumes\\"+arr[i]
                print(shutil.copyfile(original, target))
            else:
                print("Not Found")
else:
    directory = "attachments"
    dir_path = os.path.dirname(os.path.realpath(__file__))
    path = os.path.join(dir_path, directory)
    attach = os.mkdir(path)

    directory = "resumes"
    dir_path = os.path.dirname(os.path.realpath(__file__))
    path = os.path.join(dir_path, directory)
    attach = os.mkdir(path)

    outlook = win32com.client.Dispatch('outlook.application')
    mapi = outlook.GetNamespace("MAPI")
    for account in mapi.Accounts:
        print(account.DeliveryStore.DisplayName)
    inbox = mapi.GetDefaultFolder(6)
    messages = inbox.Items
    #Let's assume we want to save the email attachment to the below directory
    outputDir = r""+os.getcwd()+"\\attachments"
    try:
        for message in list(messages):
            try:
                s = message.sender
                for attachment in message.Attachments:
                    attachment.SaveASFile(os.path.join(outputDir, attachment.FileName))
                    print(f"attachment {attachment.FileName} from {s} saved")
                    
            except Exception as e:
                print("error when saving the attachment:" + str(e))
    except Exception as e:
            print("error when processing emails messages:" + str(e))

            
    arr = os.listdir(r""+os.getcwd()+"\\attachments")
    for i in range(len(arr)):
        if(('.docx') in arr[i] or ('.doc') in arr[i] or ('.pdf') in arr[i]  ):
            original = r""+os.getcwd()+"\\attachments\\"+arr[i]
            target = r""+os.getcwd()+"\\resumes\\"+arr[i]
            print(shutil.copyfile(original, target))
        else:
            print("Not Found")
    #-----------------------Doc ,  pdf sorting--------------------------------
