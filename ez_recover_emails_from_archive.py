import win32com.client
from win32com.client import constants

olFolderInbox = 6
olMailItem = 0

outlook=win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# inbox = outlook.GetDefaultFolder(6)
# inbox = outlook.GetDefaultFolder(win32com.client.constants.olFolderInbox)

# messages=inbox.Items

root_folder = outlook.Folders.Item(1)
# to get to folder
folder_2b_sorted = root_folder.Folders['2b_sorted']
# folder_2b_sorted = root_folder.Folders['Matt']
print(folder_2b_sorted)

folder_target = root_folder.Folders['Matt']

messages = folder_2b_sorted.Items

print(f'{len(messages)} found')

for message in messages:
    try:
        if "Matthew" in message.Sender.Name:
            print("-"*70)
            print(message.Subject)
            message.Move(folder_target)
    except AttributeError as err:
        print(err)