import win32com.client
from win32com.client import constants

sender_name = "Slack"
target_folder = 'Slack'

outlook=win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# inbox = outlook.GetDefaultFolder(6)
inbox = outlook.GetDefaultFolder(win32com.client.constants.olFolderInbox)

messages=inbox.Items

root_folder = outlook.Folders.Item(1)
# to get to Cloud_Automation folder
folder_target = root_folder.Folders[target_folder]

for message in messages:
    # print(message.Sender.Address)
    try:
        if sender_name in message.Sender.Name:
            print("-"*70)
            print(f'.Subject: {message.Subject}')
            message.Move(folder_target)
    except AttributeError as attrerr:
        print('. a meeting:')
        print(f'.Subject: {message.Subject}')
        print(f'.SenderName: {message.SenderName}')