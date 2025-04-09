# to be ran using "streamlit run deletion.py"
# be sure to do " pip install pywin32 streamlit" before running this
# this will only work on windows devices

import streamlit as st
import datetime as dt
import win32com.client as win32
import pythoncom

pythoncom.CoInitialize()

st.title("Deletion")

st.write("""
**Deletion** removes Outlook mail that matches your criteria.  
Specify the **mailbox**, an optional **folder path** (e.g. `Inbox/Automation/03.Spam_Verdict`),
and at least one of **Subject** or **Sender e‑mail**.  
""")

# ────────────────────────── UI ──────────────────────────
mailbox      = st.text_input("Mailbox (required)")
folder_path  = st.text_input("Folder path inside mailbox (default = Inbox)",
                             value="Inbox")

colA, colB   = st.columns(2)
with colA:
    subj_filter = st.text_input("Subject contains …", value="")
with colB:
    sender_filter = st.text_input("Sender e‑mail contains …", value="")

use_dates = st.checkbox("Limit by received‑date range")
if use_dates:
    c1, c2 = st.columns(2)
    with c1:
        start_date = st.date_input("Start", dt.date.today())
    with c2:
        end_date   = st.date_input("End",   dt.date.today())

go = st.button("Delete matching mail")

# ───────────────────── deletion logic ───────────────────
if go:

    if not mailbox:
        st.error("Mailbox is required."); st.stop()
    if not (subj_filter or sender_filter):
        st.error("Provide at least a Subject or Sender filter."); st.stop()

    try:
        outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

        # ── resolve mailbox
        rcp = outlook.CreateRecipient(mailbox)
        if not rcp.Resolve():
            st.error("Mailbox not found or you lack permission."); st.stop()

        # ── walk down the folder path
        folder_names = [f for f in folder_path.replace("\\", "/").split("/") if f]
        current_fld  = outlook.GetSharedDefaultFolder(rcp, 6)  # Inbox
        for name in folder_names[1:]:                          # skip first ("Inbox")
            current_fld = current_fld.Folders(name)

        items = current_fld.Items
        items.Sort("[ReceivedTime]", True)  # newest first

        # ── build Outlook Restrict string
        restrictions = []
        if subj_filter:
            restrictions.append(f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{subj_filter}%'")
        if sender_filter:
            restrictions.append(f"@SQL=\"urn:schemas:httpmail:fromemail\" LIKE '%{sender_filter}%'")
        if use_dates:
            start = dt.datetime.combine(start_date, dt.time.min).strftime("%m/%d/%Y %H:%M %p")
            end   = dt.datetime.combine(end_date,   dt.time.max).strftime("%m/%d/%Y %H:%M %p")
            restrictions.append(f"[ReceivedTime] >= '{start}' AND [ReceivedTime] <= '{end}'")

        if restrictions:
            restrict_str = " AND ".join(restrictions)
            items = items.Restrict(restrict_str)

        match_count = len(items)
        st.write(f"Matched **{match_count}** item(s). Deleting…")

        # ── delete
        for msg in list(items):   # cast to list to avoid iterator issues while deleting
            msg.Delete()

        st.success(f"Deletion complete – {match_count} email(s) removed "
                   f"from {mailbox} / {folder_path}")

    except Exception as e:
        st.error(f"Error: {e}")

pythoncom.CoUninitialize()