import win32com.client

outlook=win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

print(win32com.client.constants.olFolderInbox)
print(win32com.client.constants.olMailItem)
print(win32com.client.constants.olMeeting)
print(win32com.client.constants.olFolderCalendar)