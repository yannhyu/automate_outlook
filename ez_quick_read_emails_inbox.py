import win32com.client
from win32com.client import constants

outlook=win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# inbox = outlook.GetDefaultFolder(6)
inbox = outlook.GetDefaultFolder(win32com.client.constants.olFolderInbox)

messages=inbox.Items

for message in messages:
    try:
        print("-"*70)
        print(f'.Sender Name: {message.Sender.Name}')
        print(f'.Sender email: {message.Sender.Address}')
        print(" "*40)
        print(f'.RE: {message.Subject}')
        # print(f'Body: ... {message.Body[:150]}')
        print(f'Body: ... {" ".join(message.Body[:120].split())}')
        message.Unread = False

        # print(message.Body)
        # attachments = message.attachments
        # for attachment in attachments:
        #     pass
    except AttributeError as attrerr:
        # print(attrerr)
        print('. a meeting:')
        print(f'.Subject: {message.Subject}')
        print(f'.SenderName: {message.SenderName}')
        # print(f'.Start: {message.Start}')
        # print(f'.End: {message.End}')
        # print(f'.Session: {message.Session}')
        # print(f'.Body: {message.Body}')
        # print(f'.Recipients: {message.Recipients}')
