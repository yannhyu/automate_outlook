 
import win32com.client  
  
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")  
# inbox = outlook.Folders["testnotifications"].Folders["Inbox"]  
  
# messages = inbox.Items  
# for message in messages:  
#     print("Deleting a message %s" % message.Subject )  
#     message.Delete()  
# deleted = outlook.Folders["testnotifications"].Folders["Deleted Items"]  

root_folder = outlook.Folders.Item(1)
deleted = root_folder.Folders['Deleted Items']
  
while True:  
    message = deleted.Items.GetLast()  
    print("Deleting a message %s" % message.Subject)  
    message.Delete()  
print("Done")  