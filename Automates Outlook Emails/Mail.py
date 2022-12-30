import win32com.client
import schedule
import time
import os
import glob

def sends_mail():
    ol = win32com.client.Dispatch("outlook.application") 
    olmailitem = 0x0 #size of the new email
    newmail = ol.CreateItem(olmailitem) 
    newmail.Subject = 'Testing Mail' #This is where you write the subject of the email
    newmail.To = 'agen.thananchayan@signify.com; reggie.thekkumthara@signify.com;' #For more emails needed just add another email and then write a semi colon do not touch anything else 
    #agen.thananchayan@signify.com; agenchayan@gmail.com; 
    # Write email message on this line  code body message all from
    newmail.Body = 'Hello, \n\nthis is a test email to showcase how to send emails from Python and Outlook.'   #each \n means a added space on outlook so use this as a way to fomrat how you would like to send your body email
    
    attachment = r"C:\Users\670274577\OneDrive - Signify\Work\Orders\*" #copy and paste C to s from directory don't change anyhting else 

    list_of_files = glob.glob(attachment)
    latest_file = max(list_of_files, key = os.path.getctime)

    newmail.Attachments.Add(latest_file)
    # To display the mail before sending it
    # newmail.Display()
    newmail.Send()


#print('Hello, \n\nthis is a test email to showcase how to send emails from Python and Outlook.\n')
#Automates the entire proccess
schedule.every().friday.at("13:20").do(lambda: sends_mail()) # if you want to add more days just make anoter line of code just like this one and change the day and time. Example on next line
#schedule.every().thursday.at("17:44").do(lambda: sends_mail())

while True:
    schedule.run_pending()
    time.sleep(1)
