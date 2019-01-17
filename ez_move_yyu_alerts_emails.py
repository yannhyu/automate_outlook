import win32com.client
from win32com.client import constants

senders_names = ["Wickman, Albert", "Agalaba, Felix", "Yu, Yann"]
senders_emails = ["Albert.Wickman@ncr.com", "Felix.Agalaba@ncr.com", "Yann.Yu@ncr.com"]
target_folder = 'Junk Email'
source_folder = '2b_sorted'

outlook=win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

inbox = outlook.GetDefaultFolder(6)
# inbox = outlook.GetDefaultFolder(win32com.client.constants.olFolderInbox)

# messages = inbox.Items

root_folder = outlook.Folders.Item(1)

folder_source = root_folder.Folders[source_folder]
messages = folder_source.Items
# to get to Cloud_Automation folder
folder_target = root_folder.Folders[target_folder]

for message in messages:
    try:
        if any(item in message.Sender.Name for item in senders_names):
        # if sender_email == message.Sender.Address:
            print(message.Sender.Address)
            print("... moving ...")
            print("-"*70)
            print(f'.Subject: {message.Subject}')
            message.Move(folder_target)
    except AttributeError as attrerr:
        print('. a meeting:')
        # print(f'.Subject: {message.Subject}')
        # print(f'.SenderName: {message.SenderName}')