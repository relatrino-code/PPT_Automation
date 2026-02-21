# Importing libraries

import os
import datetime
import win32com.client
import pythoncom
import json


def download_attachments(config_path):
    """
    Downloads Excel attachments from Outlook emails based on subject filters and date range.

    Args:
        config_path (str): Path to the JSON configuration file. The config file should contain
                           a key 'file_paths' with a subkey 'base_path' where files will be saved.

    Functionality:
        - Connects to the user's Outlook inbox.
        - Filters emails from the current week up to today based on the "ReceivedTime".
        - Searches for emails with the subject containing "Daily Spend Report".
        - Downloads `.xlsx` attachments from the first matching email.
        - Renames the attachments based on keywords in the filename (e.g., 'installs' or 'mae').
        - Saves the attachments to the specified download folder.

    Notes:
        - Requires Outlook to be open.
        - Uses `win32com.client` to access Outlook COM objects.
        - Only processes the first matching email found within the date range.
    """

    # Initialize COM for multithreaded environments
    pythoncom.CoInitialize()
    with open(config_path, "r") as f:
        config = json.load(f)

    # Get download directory from config
    DOWNLOAD_FOLDER = config['file_paths']['base_path']

    # Define the subject to filter emails
    TARGET_SUBJECT = "Daily Spend Report"

    try:
        # Connect to Outlook
        outlook_app = win32com.client.gencache.EnsureDispatch("Outlook.Application")
        namespace = outlook_app.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)  # 6 = Inbox
    except Exception as e:
        print("❌ Failed to connect to Outlook. Make sure Outlook is open.")
        print("Error:", e)
        return

    # Get all messages from inbox
    messages = inbox.Items

    # Define time window from the start of the week (Monday) to today
    today = datetime.date.today() + datetime.timedelta(days=1)
    start_of_week = today - datetime.timedelta(days=today.weekday())  # Monday

    print(today, start_of_week)

    # Create a filter string for Outlook Restrict query
    filter_str = f"[ReceivedTime] >= '{start_of_week.strftime('%m/%d/%Y')}' AND [ReceivedTime] <= '{today.strftime('%m/%d/%Y')}'"
    filtered_messages = messages.Restrict(filter_str)
    filtered_messages.Sort("[ReceivedTime]", True)

    print(f"📬 Filtered message count: {filtered_messages.Count}")


    # Flag to track if any attachment was downloaded
    downloaded = False

    # Loop through filtered emails
    for message in filtered_messages:
        try:
            # Only process MailItem (Class 43)
            if message.Class != 43:
                continue

            subject = message.Subject or ""
            # print("Subject:", subject)

            # Check if the target subject is in the email subject line
            if TARGET_SUBJECT.lower() in subject.lower():
                # Loop through attachments
                for attachment in message.Attachments:
                    # Get the original file extension
                    original_filename = attachment.FileName
                    _, file_extension = os.path.splitext(original_filename)

                    # Only process Excel files
                    if file_extension=='.xlsx':
                        # Rename file based on keywords in original filename
                        if 'installs' in original_filename.lower() or 'mae' not in original_filename.lower():
                            # Set your custom name
                            new_filename = f"Walmart - Mobile Installs Daily Spend Report{file_extension}"  # This keeps the correct extension
                        else:
                            # Set your custom name
                            new_filename = f"Walmart App MAE Daily Spend{file_extension}"  # This keeps the correct extension

                        # Build full save path
                        save_path = os.path.join(DOWNLOAD_FOLDER, new_filename)

                        # Save the renamed file
                        attachment.SaveAsFile(save_path)

                        # printing the file details
                        print(f"📥 Downloaded: {attachment.FileName} | Subject: {message.Subject} \nRenamed as: {new_filename}| File received on: {message.ReceivedTime.strftime('%Y-%m-%d %H:%M:%S')}")
                        downloaded = True
                # Exit loop after first match is processed
                break

        except Exception as e:
            print(f"⚠️ Error processing email: {e}")

    # If no matching email was found
    if not downloaded:
        print("ℹ️ No email with target subject found this week with attachments.")

