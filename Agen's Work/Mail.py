import win32com.client
import schedule
import time

def sends_mail():
    ol = win32com.client.Dispatch("outlook.application")
    olmailitem = 0x0 #size of the new email
    newmail = ol.CreateItem(olmailitem)
    newmail.Subject = 'Testing Mail'
    newmail.To = ''
    #newmail.CC='xyz@gmail.com'
    newmail.Body = 'Hello, this is a test email to showcase how to send emails from Python and Outlook.'

    attachment = 'Work stuff.xlsx'
    newmail.Attachments.Add(attachment)
    # To display the mail before sending it
    # newmail.Display() 
    newmail.Send()

#Automates the entire proccess
schedule.every().wednesday.at("19:56").do(sends_mail())

while True:
    schedule.run_pending()
    time.sleep(1)
