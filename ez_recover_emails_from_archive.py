import win32com.client
from win32com.client import constants

outlook=win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# inbox = outlook.GetDefaultFolder(6)
# inbox = outlook.GetDefaultFolder(win32com.client.constants.olFolderInbox)

# messages=inbox.Items

root_folder = outlook.Folders.Item(1)
# to get to folder
folder_andrew = root_folder.Folders['Andrew']

for message in messages:
    # print(message.Sender.Address)
    if "Sender Name" in message.Sender.Name:
        print("-"*70)
        print(message.Subject)
        message.Move(folder_andrew)