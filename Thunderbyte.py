import win32com.client
from pathlib import Path
import re
import hashlib
import os 
import json

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

def get_subfolders(selected_folder, folder, selected_folder_names, indent=0):
    subfolders = folder.Folders
    if folder.Name in selected_folder_names:
        selected_folder_names.remove(folder.Name)
        # Do something with the selected folder and its subfolders
        if folder == selected_folder:
            print("Selected folder:", folder.Name)
            # ...
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
        selected_folder_numbers = input("Enter the number(s), a to save or q to exit: ")
        if selected_folder_numbers.lower() == "a":
            global subfolder_location
            subfolder_location = folder.FolderPath
            break
        elif selected_folder_numbers.lower() == "q":
            break
        
        selected_folder_numbers = selected_folder_numbers.split(",")
        if len(selected_folder_numbers) > 3:
            print("Cannot select more than 3 folders.")
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

print(subfolder_location)

output_dir = Path.cwd() / "Output"
output_dir.mkdir(parents=True, exist_ok=True)

# Define a function to retrieve messages from a specific subfolder
def get_messages(folder_path):

    path = folder_path

# Split the string using the backslash as the separator
    parts = path.split("\\")

    # Extract the individual parts
    a = parts[1]
    b = parts[2]
    c = parts[3]

    folder = outlook.Folders.Item(a).Folders.Item(b).Folders.Item(c) # need to test this 
    messages = folder.Items
    message_info_list = []
    for message in messages:
        subject = message.Subject
        body = message.body # test with the data for the task
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
folder_path = subfolder_location # Modify this to select the desired subfolder
get_messages(folder_path)


