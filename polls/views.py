from django.shortcuts import render,  redirect

# Create your views here.
from django.http import HttpResponse
dotpaths = '.'

def index(request):
    return HttpResponse("Hello, world !")

import pandas as pd
from .forms import UploadFileForm
from .models import YourModel



def upload_file(request):
    if request.method == 'POST':
        form = UploadFileForm(request.POST, request.FILES)
        if form.is_valid():
            file = request.FILES['file']
            df = pd.read_excel(file)

            # Iterate over the rows of the dataframe and create model instances
            for _, row in df.iterrows():
                # YourModel.objects.create(
                #     field1=row[0],
                #     field2=row[1],
                #     field3=row[2],
                #     field4=row[3],
                #     # Add more fields as necessary
                # )
                # Check if the data already exists in the database
                if YourModel.objects.filter(field1=row[0], field2=row[1], field3=row[2], field4=row[3]).exists():
                    continue  # Skip creating a new instance
                YourModel.objects.create(
                    field1=row[0],
                    field2=row[1],
                    field3=row[2],
                    field4=row[3],
                    # Add more fields as necessary
                )


            # Optionally, you can redirect the user to a success page
            # return redirect('success')
            return render(request, 'upload.html', {'form': form})
    else:
        form = UploadFileForm()

    return render(request, 'upload.html', {'form': form})



import smtplib
import imaplib
import email
import pandas as pd
import numpy as np

def emailscraper(request):
    try:
        username = "username"
        password = "password"
        # SMTP server settings
        smtp_server = "smtp.ionos.com"
        smtp_port = 465
        smtp_username = username
        smtp_password = password

        # IMAP server settings
        imap_server = "imap.ionos.com"
        imap_port = 993
        imap_username = username
        imap_password = password


        # Output file name
        output_file = dotpaths+'/media/input/email_data.txt'

        # Connect to the IMAP server
        imap_connection = imaplib.IMAP4_SSL(imap_server, imap_port)
        imap_connection.login(imap_username, imap_password)

        # Select the INBOX mailbox
        mailbox_name = "INBOX"
        imap_connection.select(mailbox_name)

        # Search for all emails in the INBOX
        _, email_ids = imap_connection.search(None, "ALL")

        # Open the output file in write mode
        with open(output_file, "w",  encoding="utf-8", errors="ignore") as file:
            # Iterate over the email IDs in reverse order
            for email_id in reversed(email_ids[0].split()):
                _, email_data = imap_connection.fetch(email_id, "(RFC822)")

                # Parse the email data
                raw_email = email_data[0][1]
                parsed_email = email.message_from_bytes(raw_email)

                # Extract the desired email information
                subject = parsed_email["Subject"]
                sender = parsed_email["From"]
                received_time = parsed_email["Date"]

                # Get the email body
                if parsed_email.is_multipart():
                    for part in parsed_email.walk():
                        if part.get_content_type() == "text/plain":
                            email_body = part.get_payload(decode=True).decode("utf-8", errors="ignore")
                            break
                else:
                    email_body = parsed_email.get_payload(decode=True).decode("utf-8", errors="ignore")

                # Write the email information to the file
                file.write("Subject: " + subject + "\n")
                file.write("Sender: " + sender + "\n")
                file.write("Received Time: " + received_time + "\n")
                file.write("Email Body: " + email_body + "\n")
                file.write("---------------------------------------------\n")

        # Close the IMAP connection
        imap_connection.close()
        imap_connection.logout()
        # return HttpResponse("Hello, world !")
            
        # Read the text from a file
        inputtextfile = dotpaths+'/media/input/email_data.txt'
        # with open(inputtextfile, 'r') as file:
        with open(inputtextfile, encoding="utf8") as file:
            text = file.read()

        # Split the text into sections based on "Subject:-"
        sections = text.split("Subject:")

        # Extract subjects, senders, received times, and messages
        subjects = []
        senders = []
        received_times = []
        messages = []

        for section in sections[1:]:
            lines = section.strip().split("\n")
            subject = lines[0].strip()
            sender = lines[1].strip()
            received_time = lines[2].strip()
            message = "\n".join(lines[4:]).strip()

            subjects.append(subject)
            senders.append(sender)
            received_times.append(received_time)
            messages.append(message)

        # Find the maximum length among the extracted arrays
        max_length = max(len(subjects), len(senders), len(received_times), len(messages))

        # Extend the arrays to the maximum length by filling with NaN values
        subjects += [np.nan] * (max_length - len(subjects))
        senders += [np.nan] * (max_length - len(senders))
        received_times += [np.nan] * (max_length - len(received_times))
        messages += [np.nan] * (max_length - len(messages))

        # Create a DataFrame with the data
        df = pd.DataFrame({'Subject': subjects, 'Sender': senders, 'Received Time': received_times, 'Message': messages})

        # Save the DataFrame to an XLSX file
        outputexcelfile = df.to_excel(dotpaths+'/media/input/email_subjects.xlsx', index=False)
        outputexcelfilepaths = dotpaths+'/media/input/email_subjects.xlsx'
        # Display the DataFrame
        # print(df)
        #df.head(10)
        df = pd.read_excel(outputexcelfilepaths)

        # Iterate over the rows of the dataframe and create model instances
        for _, row in df.iterrows():
                    # YourModel.objects.create(
                    #     field1=row[0],
                    #     field2=row[1],
                    #     field3=row[2],
                    #     field4=row[3],
                    #     # Add more fields as necessary
                    # )
                    # Check if the data already exists in the database
                    if YourModel.objects.filter(field1=row[0], field2=row[1], field3=row[2], field4=row[3]).exists():
                        continue  # Skip creating a new instance
                    YourModel.objects.create(
                        field1=row[0],
                        field2=row[1],
                        field3=row[2],
                        field4=row[3],
                        # Add more fields as necessary
                    )
        return HttpResponse("Email data has been updated")
    except:
        return HttpResponse("Email data has not been updated")


