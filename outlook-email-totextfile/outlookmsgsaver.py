import win32com.client
import csv

def extract_email_content_and_save_csv(csv_file_path):
    # Connect to Outlook
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    # Get the default Inbox folder
    inbox_folder = namespace.GetDefaultFolder(6)  # 6 corresponds to the Inbox folder

    # Get the email items in the Inbox
    email_items = inbox_folder.Items

    # Create a CSV file
    with open(csv_file_path, 'w', newline='') as csv_file:
        writer = csv.writer(csv_file)

        # Write the header row
        writer.writerow(['Subject', 'Sender', 'Received Time', 'Body'])

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

            # Write the email content to the CSV file
            writer.writerow([subject, sender, received_time, body])

    print(f"Email content saved to CSV file: {csv_file_path}")

# Usage
csv_file_path = './email_content.csv'
extract_email_content_and_save_csv(csv_file_path)
