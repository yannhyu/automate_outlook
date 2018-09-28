import win32com.client
from win32com.client import constants

outlook=win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# inbox = outlook.GetDefaultFolder(6)
# inbox = outlook.GetDefaultFolder(win32com.client.constants.olFolderInbox)


root_folder = outlook.Folders.Item(1)
inbox = root_folder.Folders['2b_sorted']

messages=inbox.Items

for message in messages:
    # print("-"*70)
    # print(f'.Sender Name: {message.Sender.Name}')
    # print(f'.Sender email: {message.Sender.Address}')
    # print(f'.RE: {message.Subject}')
    # print(f'Body: ... {message.Body[:150]}')
    try:
        if 'Yu, Yann' in message.Sender.Name:
            print("-"*70)
            print(f'.Sender Name: {message.Sender.Name}')
            print(f'.Sender email: {message.Sender.Address}')
            print(f'.RE: {message.Subject}')
    except AttributeError as attrerr:
        print(attrerr)       
