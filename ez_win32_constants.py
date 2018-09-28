import win32com.client

outlook=win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

print(win32com.client.constants.olFolderInbox)
print(win32com.client.constants.olMailItem)
print(win32com.client.constants.olMeeting)
print(win32com.client.constants.olFolderCalendar)
print(win32com.client.constants.olFolderDeletedItems)

'''
how to reach any default folder not just "Inbox" here's the list:

3  Deleted Items
4  Outbox
5  Sent Items
6  Inbox
9  Calendar
10 Contacts
11 Journal
12 Notes
13 Tasks
14 Drafts
'''