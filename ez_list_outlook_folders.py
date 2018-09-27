from win32com.client import Dispatch

outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")
root_folder = outlook.Folders.Item(1)
print(root_folder.Name)

# to know the names of the subfolders:
for folder in root_folder.Folders:
    print(folder.Name)

# to get to 2b_sorted folder
folder_2b_sorted = root_folder.Folders['2b_sorted']
print(folder_2b_sorted.Name)  