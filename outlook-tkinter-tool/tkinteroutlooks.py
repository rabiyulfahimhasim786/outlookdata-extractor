import win32com.client
import tkinter as tk
import webbrowser

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
            # Check if the subject matches the input value
            if email.Subject == subject_entry.get():
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

    output_label.config(text="Email content extracted and saved to:")
    output_link.config(state=tk.NORMAL)

def open_output_file():
    webbrowser.open("email_content.txt")

# Tkinter GUI
window = tk.Tk()
window.title("Email Extraction")
window.geometry("300x200")

# Subject Entry
subject_label = tk.Label(window, text="Enter Subject:")
subject_label.pack()
subject_entry = tk.Entry(window)
subject_entry.pack()

# Button
extract_button = tk.Button(window, text="Extract", command=extract_email_content)
extract_button.pack()

# Output Link
output_label = tk.Label(window, text="")
output_label.pack()
output_link = tk.Button(window, text="Open Output File", state=tk.DISABLED, command=open_output_file)
output_link.pack()

window.mainloop()
