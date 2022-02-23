###### PRE-REQUISITES####
# You are expected to set the Outlook applicatoion to "Work Offline" mode before starting the mail merge. 
# The code assumes that you are using the default Mail merge feature of Microsoft word to generate email messages using email list from which ever source feasible
# This way you need not create your HTML template separately. You can apply all fonts and formatting on word and create your email body as required. 
# This code will only manipulate the From / Sender's address and work around the email sending limits of 30 mails per minute.

from datetime import datetime
import win32com.client as client
from time import sleep

startTime = datetime.now()
outlook = client.Dispatch('Outlook.Application')
namespace = outlook.GetNameSpace('MAPI')

# Get details of your Default Mail Account setup on Outlook
mail_account = input("Enter the default mail account that contains your Drafts and Outbox folders: " )

# Get details of the Mail Account you want to use, to send emails.
onbehalfof = input("Enter the Email ID you want to send on behalf of: ")

user = namespace.Folders[mail_account]
outbox = user.Folders['Outbox']
drafts = user.Folders['Drafts']

messages = [message for message in outbox.Items]

for message in messages:
    message.SentOnBehalfOfName = onbehalfof
    message.Move(drafts)

drafted = [email for email in drafts.Items]
chunks = [drafted[x:x+30] for x in range(0, len(drafted), 30)]


def go_online():
    
    if namespace.Offline:
        print('Outlook is currently set to "Work Offline"')
        
    done = input('Switch Outlook to work online mode and enter YES or DONE: ')
    return done


def send_chunks(done):
    
    if done.upper() in ['YES', 'S', 'DONE'] and not (namespace.Offline):
        for chunk in chunks:
            # iterate through each recipient in chunk and send mail
            for email in chunk:
                email.Display()
                email.Send()
            # wait 60 seconds before sending next chunk
            sleep(60)
    else:
        done = go_online()
        send_chunks(done)
        

send_chunks('No')
lapse = datetime.now()-startTime
print('Time taken: ',lapse)

