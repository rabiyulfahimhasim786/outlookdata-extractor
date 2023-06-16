import win32com.client
import pandas as pd

def extract_email_content():
    # Connect to Outlook
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    # Get the default Inbox folder
    inbox_folder = namespace.GetDefaultFolder(6)  # 6 corresponds to the Inbox folder

    # Get the email items in the Inbox
    email_items = inbox_folder.Items

    email_content = []

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

        # Add email content to the list
        email_content.append({
            'Subject': subject,
            'Sender': sender,
            'Received Time': received_time,
            'Body': body
        })

    return email_content


email_content = extract_email_content()
df = pd.DataFrame(email_content)

csv_file_path = './email_contents.csv'
df.to_csv(csv_file_path, index=False)



