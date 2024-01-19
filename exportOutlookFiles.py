from pathlib import Path;
import datetime;
import re;
import win32com.client;

# Create output folders for emails with and without "invoice"
output_dir_invoice = Path.cwd() / "WithInvoice"
output_dir_no_invoice = Path.cwd() / "WithoutInvoice"
output_dir_invoice.mkdir(parents=True, exist_ok=True)
output_dir_no_invoice.mkdir(parents=True, exist_ok=True)

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")



# Get Outlook folder(inbox) for processed emails
inbox = outlook.GetDefaultFolder(6)


# Docs    => https://docs.microsoft.com/en-us/office/vba/api/outlook.oldefaultfolders
# Folders => DeletedItems=3, Outbox=4, SentMail=5, Inbox=6, Drafts=16, FolderJunk=23

messages = inbox.Items

for message in messages:
    subject = message.Subject
    body = message.body
    attachments = message.Attachments

    # Check if the word "invoice" is present in the subject or body
    if "invoice" in subject.lower() or "invoice" in body.lower() or "invoice" in attachments.FileName:
        target_folder = output_dir_invoice / re.sub('[^0-9a-zA-Z]+', '', subject)
    else:
        target_folder = output_dir_no_invoice / re.sub('[^0-9a-zA-Z]+', '', subject)

    # Create separate folder for each message, exclude special characters and timestamp
    current_time = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    target_folder = target_folder / current_time
    target_folder.mkdir(parents=True, exist_ok=True)

    # Write body to text file
    Path(target_folder / "EMAIL_BODY.txt").write_text(str(body))

    # Save attachments and exclude special characters
    for attachment in attachments:
        filename = re.sub('[^0-9a-zA-Z\.]+', '', attachment.FileName)
        attachment.SaveAsFile( target_folder / filename)

