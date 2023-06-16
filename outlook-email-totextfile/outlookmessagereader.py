import win32com.client

def extract_email_content():
    # Connect to Outlook
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    # Get the default Inbox folder
    inbox_folder = namespace.GetDefaultFolder(6)  # 6 corresponds to the Inbox folder

    # Get the email items in the Inbox
    email_items = inbox_folder.Items

    # Iterate through the email items
    for email in email_items:
        # Extract email properties
        subject = email.Subject
        sender = email.SenderEmailAddress
        received_time = email.ReceivedTime

        # Extract email body
        if email.BodyFormat == 2:  # Plain Text format
            body = email.Body
        elif email.BodyFormat == 3:  # HTML format
            body = email.HTMLBody

        # Print or process the extracted email content
        print("Subject:", subject)
        print("Sender:", sender)
        print("Received Time:", received_time)
        print("Body:", body)
        print("---")

# Usage
extract_email_content()
