import win32com.client
from win32com.client import constants

outlook=win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# inbox = outlook.GetDefaultFolder(6)
# inbox = outlook.GetDefaultFolder(win32com.client.constants.olFolderInbox)


root_folder = outlook.Folders.Item(1)
folder_deleted_items = root_folder.Folders['Deleted Items']

messages=folder_deleted_items.Items

print(f'Found {len(messages)} in Deleted Items folder')
for message in messages:
    try:
        print("-"*70)
        print(f'.Sender Name: {message.Sender.Name}')
        print(f'.Sender email: {message.Sender.Address}')
        print(f'.RE: {message.Subject}')
        print(f'Body: ... {message.Body[:150]}')

        if 'Yu, Yann' in message.Sender.Name:
            print(f'... deleting {message.Subject}')
            message.Delete()        
    except AttributeError as attrerr:
        print(attrerr)     
