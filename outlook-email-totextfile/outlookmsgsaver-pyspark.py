import win32com.client
from pyspark.sql import SparkSession

spark = SparkSession.builder.appName("EmailContent").getOrCreate()

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
        email_content.append((subject, sender, received_time, body))

    return email_content

email_content = extract_email_content()
email_rdd = spark.sparkContext.parallelize(email_content)
df = spark.createDataFrame(email_rdd, ["Subject", "Sender", "Received Time", "Body"])

csv_file_path = './email_content.csv'
df.write.csv(csv_file_path, header=True)
