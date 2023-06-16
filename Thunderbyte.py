import win32com.client
from pathlib import Path
import re
import hashlib
import os
import json
import win32ui

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
# Checking for PST loaded into outlook

def outlook_running():
    try:
        win32ui.FindWindow(None, "Microsoft Outlook")
        return True
    except win32ui.error:
        return False

if not outlook_running():
    os.startfile("outlook")
    
def check_pst_loaded():
    for store in outlook.Stores:
        if store.ExchangeStoreType == 3:  # 3 indicates PST store
            return True

    return False

if check_pst_loaded():
    print("PST found")
    print("Continuing........")
elif not check_pst_loaded():
    print("Please load PST into outlook")
    exit()
else:
    print("Unable to read outlook")
    exit()

def get_subfolders(selected_folder, folder, selected_folder_names, indent=0):
    subfolders = folder.Folders
    if folder.Name in selected_folder_names:
        selected_folder_names.remove(folder.Name)
        # Do something with the selected folder and its subfolders
        if folder == selected_folder:
            print("Selected folder:", folder.Name)
            if len(selected_folder_names) == 0:
                return
            
    if len(subfolders) > 0:
        for subfolder in subfolders:
            print(" " * indent + subfolder.Name)
            get_subfolders(selected_folder, subfolder, selected_folder_names.copy(), indent+4)

def print_all_folders():
    top_level_folders = outlook.Folders
    selected_folders = []
    while True:
        if len(selected_folders) == 0:
            print("Available folders:")
            for i, folder in enumerate(top_level_folders):
                print(f"{i+1}. {folder.Name}")
        else:
            print("Available subfolders:")
            for i, folder in enumerate(selected_folders[-1].Folders):
                print(f"{i+1}. {folder.Name}")
        selected_folder_numbers = input("Enter the number(s), s to save or q to exit: ")
        if selected_folder_numbers.lower() == "s":
            global subfolder_location
            subfolder_location = folder.FolderPath
            break
        elif selected_folder_numbers.lower() == "q":
            break
        
        selected_folder_numbers = selected_folder_numbers.split(",")
        if len(selected_folder_numbers) > 5: # test to see limit of script
            print("Cannot select more than 5 folders.")
            continue
        selected_folders_copy = selected_folders.copy()
        for folder_number in selected_folder_numbers:
            try:
                folder_number = int(folder_number)
            except ValueError:
                print("Invalid input:", folder_number)
                selected_folders = selected_folders_copy
                break
            if len(selected_folders) == 0:
                if folder_number < 1 or folder_number > len(top_level_folders):
                    print("Invalid folder number:", folder_number)
                    selected_folders = selected_folders_copy
                    break
                selected_folders.append(top_level_folders[folder_number-1])
            else:
                subfolders = selected_folders[-1].Folders
                if folder_number < 1 or folder_number > len(subfolders):
                    print("Invalid folder number:", folder_number)
                    selected_folders = selected_folders_copy
                    break
                selected_folders.append(subfolders[folder_number-1])
        else:
            selected_folder = selected_folders[-1]
            selected_folder_names = [selected_folder.Name]
            get_subfolders(selected_folder, selected_folder, selected_folder_names, indent=4)

print_all_folders()

## this is to be used for later scripts 
# Added in to combat the error if "q" pressed to exit
if "subfolder_location" in locals():
    print(subfolder_location)
# This get the current working dir and make a folder called output
output_dir = Path.cwd() / "Output"
output_dir.mkdir(parents=True, exist_ok=True)

# Define a function to retrieve messages from a specific subfolder
def get_messages(folder_path):
    path = folder_path
    parts = path.split("\\")

    # Retrieve the top-level folder
    top_folder_name = parts[1]
    folder = outlook.Folders.Item(top_folder_name)

    # Traverse the remaining folder hierarchy
    # This should read the full "folder" path from the outlook
    for folder_name in parts[2:]:
        subfolders = folder.Folders
        folder = subfolders.Item(folder_name)

    # Retrieve the messages from the final folder
    messages = folder.Items
    all_message_info = []  # List to store all message metadata
    
    def retrieve_msg_attachment_metadata(msg_attachment_path):
        # Implement your logic to retrieve metadata from .msg attachments
        # and return a dictionary containing the metadata
        metadata = {
            "subject": "Attachment Subject",
            "body": "Attachment Body",
            "attachments": []
        }
        return metadata

    for message in messages:
        subject = message.Subject
        body = message.body

        # Create separate folder for each message, exclude special characters
        target_folder = output_dir / re.sub('[^0-9a-zA-Z]+', '', subject)
        target_folder.mkdir(parents=True, exist_ok=True)

        # Create subfolder for attachments
        attachments_folder = target_folder / 'attachments'
        attachments_folder.mkdir(parents=True, exist_ok=True)

        attachments = message.Attachments
        attachment_info_list = []

        for attachment in attachments:
            filename = re.sub('[^0-9a-zA-Z\.]+', '', attachment.FileName)
            attachment_path = attachments_folder / filename

            if attachment_path.suffix.lower() == ".msg":
                attachment_metadata = retrieve_msg_attachment_metadata(attachment_path)
                if attachment_metadata is None:
                    attachment_info = {
                        "filename": filename,
                        "hash": "",
                        "subject": "",
                        "body": "",
                        "attachments": []
                    }
                else:
                    attachment_info = {
                        "filename": filename,
                        "hash": attachment_metadata.get("hash", ""),
                        "subject": attachment_metadata.get("subject", ""),
                        "body": attachment_metadata.get("body", ""),
                        "attachments": attachment_metadata.get("attachments", [])
                    }
            else:
                try:
                    attachment.SaveAsFile(str(attachment_path))
                except:
                    pass

                with open(attachment_path, "rb") as f:
                    attachment_hash = hashlib.sha256(f.read()).hexdigest()

                attachment_info = {
                    "filename": filename,
                    "hash": attachment_hash
                }
            attachment_info_list.append(attachment_info)

        message_info = {
            "subject": subject,
            "body": str(body),
            "attachments": attachment_info_list
        }
        all_message_info.append(message_info)

    # Create the parent dictionary with the "emails" key and value as the list of message metadata
    parent_dict = {
        "emails": all_message_info
    }

    # Save all message metadata to a single JSON file
    with open(output_dir / "all_messages.json", "w") as json_file:
        json.dump(parent_dict, json_file)

# Added Error handling

try:
    #This should fix the error if the script is exited before its complete
    folder_path = subfolder_location
    get_messages(folder_path)
except NameError:
    print("Exiting....")
