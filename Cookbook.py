import win32com.client  #pip install pywin32
from pathlib import Path
import re
import hashlib
import os 
import json

# Connect to outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

output_dir = Path.cwd() / "Output"
output_dir.mkdir(parents=True, exist_ok=True)

# Define a function to retrieve messages from a specific subfolder
def get_messages(folder_path):
    folder = outlook.Folders.Item("Outlook Data File").Folders.Item("2023").Folders.Item("2023-01") # Modify this line to select the desired subfolder
    messages = folder.Items
    message_info_list = []
    for message in messages:
        subject = message.Subject
        body = message.body
        attachments = message.Attachments

        # Create separate folder for each message, exclude special characters
        target_folder = output_dir / re.sub('[^0-9a-zA-Z]+', '', subject) 
        target_folder.mkdir(parents=True, exist_ok=True)

        # Create subfolder for attachments
        attachments_folder = target_folder / 'attachments'
        attachments_folder.mkdir(parents=True, exist_ok=True)

        # Save attachments and exclude special characters in the filename
        attachment_info_list = []
        for attachment in attachments:
            filename = re.sub('[^0-9a-zA-Z\.]+', '', attachment.FileName)
            try:
                attachment.SaveAsFile(attachments_folder / filename)
            except:
                pass
            
            # Calculate attachment hash
            with open(attachments_folder / filename, 'rb') as f:
                attachment_hash = hashlib.sha256(f.read()).hexdigest()

            # Save attachment hash to JSON file
            attachment_info = {
                "filename": filename,
                "hash": attachment_hash
            }
            attachment_info_list.append(attachment_info)

            # Remove the original attachment file
            os.remove(attachments_folder / filename)

        # Save message metadata to JSON file, including body of the email
        message_info = {
            "subject": subject,
            "body": str(body),
            "attachments": attachment_info_list
        }
        message_info_list.append(message_info)
        with open(target_folder / "message_info.json", "w") as f:
            json.dump(message_info, f)

    # Save all message metadata to a single JSON file
    with open(output_dir / f"{re.sub('[^0-9a-zA-Z]+', '', folder.Name)}.json", "w") as f:
        json.dump(message_info_list, f)

# Retrieve messages from the specified subfolder
folder_path = r"\\Outlook Data File\2023\2023-01" # Modify this to select the desired subfolder
get_messages(folder_path)
