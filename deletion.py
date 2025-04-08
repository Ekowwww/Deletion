# to be ran using "streamlit run deletion.py"
# be sure to do " pip install pywin32 streamlit" before running this
# this will only work on windows devices

import streamlit as st
import datetime
import win32com.client as win32

st.title("Deletion")

st.write(
    """
    **Deletion** is a tool to help you remove unwanted emails from Outlook 
    by specifying the **mailbox**, the email's **sender**, or its **subject**. 
    You can optionally provide a date range to limit which emails are deleted.

    **Warning**: Deletions are permanent! Use with caution.
    """
)

# 1) Mailbox to target (required)
mailbox_input = st.text_input("Mailbox (required)", value="")
st.caption("Enter the email address or display name of the mailbox you have permission to access.")

# 2) Subject filter (required)
subject_filter = st.text_input("Filter by Subject (required)", value="")

# 3) Sender filter (required)
sender_filter = st.text_input("Filter by Sender (required)", value="")

# 4) Optional date range
use_date_filter = st.checkbox("Filter by Date Range?")
start_date = None
end_date = None
if use_date_filter:
    col1, col2 = st.columns(2)
    with col1:
        start_date = st.date_input("Start Date", datetime.date.today())
    with col2:
        end_date = st.date_input("End Date", datetime.date.today())

# 5) Deletion button
if st.button("Delete Matching Emails"):
    # Validate required fields
    if not mailbox_input:
        st.error("Mailbox is required.")
        st.stop()

    if not subject_filter and not sender_filter:
        st.error("You must provide at least a subject or a sender filter.")
        st.stop()

    try:
        # Connect to Outlook
        outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

        # Resolve mailbox
        recipient = outlook.CreateRecipient(mailbox_input)
        if not recipient.Resolve():
            st.error("Could not find or resolve the specified mailbox. "
                     "Please check the name/email and your permissions.")
            st.stop()

        # Access the Inbox of the specified mailbox (Folder type 6 = Inbox)
        mailbox_inbox = outlook.GetSharedDefaultFolder(recipient, 6)
        messages = mailbox_inbox.Items

        # Build a list of messages to delete
        to_delete = []

        for message in messages:
            # Check subject & sender
            subject_matches = (subject_filter.lower() in message.Subject.lower()) \
                if subject_filter else False
            sender_matches = (sender_filter.lower() in str(message.SenderName).lower()) \
                if sender_filter else False

            if subject_matches and sender_matches:
                if use_date_filter and start_date and end_date:
                    received_time = message.ReceivedTime
                    # pywin32 usually provides datetime directly, but let's ensure it's a Python datetime
                    received_time = datetime.datetime.fromtimestamp(int(received_time.timestamp()))
                    
                    if start_date <= received_time.date() <= end_date:
                        to_delete.append(message)
                else:
                    to_delete.append(message)

        # Perform the deletion
        deleted_count = 0
        for msg in to_delete:
            msg.Delete()
            deleted_count += 1

        st.success(f"Deletion complete. {deleted_count} matching email(s) deleted from '{mailbox_input}'.")
    except Exception as e:
        st.error(f"An error occurred: {e}")