import win32com.client
import schedule
import time
import os
import glob

def sends_mail():
    ol = win32com.client.Dispatch("outlook.application")
    olmailitem = 0x0 #size of the new email
    newmail = ol.CreateItem(olmailitem)
    newmail.Subject = 'Testing Mail'
    newmail.To = 'Vithursh0512@gmail.com; Vithuchayan@gmail.com; agenchayan@gmail.com; st.tc270@gmail.com;'
    newmail.Body = 'Hello, this is a test email to showcase how to send emails from Python and Outlook.'

    attachment = r"C:\Users\vithu\OneDrive\Desktop\Agen's Work\Data\*"

    list_of_files = glob.glob(attachment)
    latest_file = max(list_of_files, key = os.path.getctime)

    newmail.Attachments.Add(latest_file)
    # To display the mail before sending it
    # newmail.Display()
    newmail.Send()

#Automates the entire proccess
schedule.every().thursday.at("17:57").do(lambda: sends_mail())
#schedule.every().thursday.at("17:44").do(lambda: sends_mail())

while True:
    schedule.run_pending()
    time.sleep(1)
