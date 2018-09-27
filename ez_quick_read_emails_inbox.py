import win32com.client
from win32com.client import constants

outlook=win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# inbox = outlook.GetDefaultFolder(6)
inbox = outlook.GetDefaultFolder(win32com.client.constants.olFolderInbox)

messages=inbox.Items

for message in messages:
    print("-"*70)
    print(message.Subject)
    print(f'... {message.Body[:150]}')
    # print(message.Body)

    # attachments = message.attachments

    # for attachment in attachments:
    #     pass