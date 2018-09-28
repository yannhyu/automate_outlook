#!/usr/bin/env python
# -*- coding: utf-8 -*-
import os, sys
import win32com.client
from datetime import datetime
from datetime import timedelta
import pywintypes
import pytz

utc=pytz.UTC

# Microsoft Outlook Constants
# http://msdn.microsoft.com/en-us/library/aa219371(office.11).aspx
olFolderDeletedItems=3
olFolderSentMail=5
# or you can use the following command to generate
# c:\python26\python.exe c:\Python26\lib\site-packages\win32com\client\makepy.py -d
# After generated, you can use win32com.client.constants.olFolderSentMail

# http://code.activestate.com/recipes/496683-converting-ole-datetime-values-into-python-datetim/
OLE_TIME_ZERO = datetime(1899, 12, 30, 0, 0, 0)
def ole2datetime(oledt):
    return OLE_TIME_ZERO + timedelta(days=float(oledt))

if __name__ == '__main__':
    app = win32com.client.Dispatch( "Outlook.Application" )
    ns = app.GetNamespace( "MAPI" )
    folders = [
        #ns.GetDefaultFolder(olFolderSentMail),
        ns.GetDefaultFolder(olFolderDeletedItems)
    ]
    for folder in folders:
        print( "Processing %s" % folder.Name )

        past30days=datetime.now()-timedelta(days=30)
        mark2delete=[]
        #If you use makepy.py, you have to use the following codes instead of "for item in folder.Items"
        #for i in range(1,folder.Items.Count+1):
        #    item = folder.Items[i]
        for item in folder.Items:
            # if ole2datetime(item.LastModificationTime)<past30days:
            if item.LastModificationTime.replace(tzinfo=None) < past30days:
                mark2delete.append( item )
        if len(mark2delete)>0:
            for item in mark2delete:
                print( "Removing %s" % item.Subject )
                item.Delete()
        else:
            print("No matched mails.")