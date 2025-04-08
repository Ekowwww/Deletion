How This Works
Streamlit Inputs:

Subject Filter: We check whether the user-provided string appears in the message’s 
Subject.

Sender Filter: We check whether the user-provided string appears in the message’s SenderName.
Both are required fields in this example to demonstrate the usage. Adjust the logic if you want to allow only one field.

Date Range (Optional):
If the user checks “Filter by Date Range?”, the script will only delete emails whose ReceivedTime falls between the 
selected start and end dates.

Outlook Deletion:
The code uses win32com.client.Dispatch("Outlook.Application") to get the MAPI namespace and then retrieves the 
Inbox with GetDefaultFolder(6). We then iterate over all items in the Inbox, check the filters, and call Delete() on matching messages.

Permanent Deletion:
Outlook’s Delete() moves the emails to the Deleted Items folder. Depending on your Outlook settings, they may be 
permanently gone if that folder is emptied automatically. Always confirm and use caution when running this!