import win32com.client

def extract_email_content():
    # Connect to Outlook
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    # Get the default Inbox folder
    inbox_folder = namespace.GetDefaultFolder(6)  # 6 corresponds to the Inbox folder

    # Get the email items in the Inbox
    email_items = inbox_folder.Items

    # Create a text file to save the extracted email content
    with open("email_content.txt", "w", encoding="utf-8") as file:
        # Iterate through the email items
        for email in email_items:
            # Check if the subject matches "My Hotlist"
            if email.Subject == "FW: My Hotlist":
                # Extract email properties
                subject = email.Subject
                sender = email.SenderEmailAddress
                received_time = email.ReceivedTime

                # Extract email body
                if email.BodyFormat == 2:  # Plain Text format
                    body = email.Body
                elif email.BodyFormat == 3:  # HTML format
                    body = email.HTMLBody

                 # Remove unwanted empty lines from each line
                body_lines = body.splitlines()
                body_lines = [line for line in body_lines if line.strip() != '']
                body = '\n'.join(body_lines)

                # Write the extracted email content to the text file
                file.write("Subject: {}\n".format(subject))
                file.write("Sender: {}\n".format(sender))
                file.write("Received Time: {}\n".format(received_time))
                file.write("Body: {}\n".format(body))
                file.write("---\n")

    print("Email content extracted and saved to email_content.txt")

# Usage
extract_email_content()
