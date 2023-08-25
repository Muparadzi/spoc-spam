import os
import json
from datetime import datetime
import hashlib
import fnmatch
import email
from extract_msg import Message
import subprocess
import calendar


def calculate_hash(file_path):
    hash_object = hashlib.sha256()

    with open(file_path, 'rb') as f:
        while chunk := f.read(8192):
            hash_object.update(chunk)

    return hash_object.hexdigest()


def extract_msg_attachment_info(msg_attachment):
    msg = Message(msg_attachment)
    subject = msg.subject
    hash_value = calculate_hash(msg_attachment)
    return {
        "Attachment Name": os.path.basename(msg_attachment),
        "Attachment Hash": hash_value,
        "Subject": subject
    }

def extract_email_info(email_message, file_path):
    sender = email_message.get("From", "Unknown Sender")
    recipient = email_message.get("To", "Unknown Recipient")
    subject = email_message.get("Subject", "No Subject")
    date_str = email_message.get("Date", "")
    Message_Id = email_message.get("Message-Id", "")


    date_obj = email.utils.parsedate_to_datetime(date_str)
    formatted_date = date_obj.strftime("%Y-%m-%d %H:%M:%S")

    hash_value = calculate_hash(file_path)
    file_name = os.path.basename(file_path)

    # Extract attachment information
    attachments = []
    for part in email_message.walk():
        if part.get_content_maintype() == "application" and part.get_content_disposition() == "attachment":
            attachment_name = part.get_filename()
            attachment_data = part.get_payload(decode=True)

            # If attachment is a .msg file
            if attachment_name.lower().endswith(".msg"):
                msg_attachment_info = extract_msg_attachment_info(attachment_data)
                attachments.append(msg_attachment_info)
            else:
                attachment_hash = calculate_hash(attachment_data)
                attachments.append({
                    "Attachment Name": attachment_name,
                    "Attachment Hash": attachment_hash
                })

    return {
        "Sender": sender,
        "Recipient": recipient,
        "Subject": subject,
        "Date": formatted_date,
        "Hash": hash_value,
        "Message Id" : Message_Id,
        "File Name": file_name,
        "Attachments": attachments
    }

def extract_emails_from_pst(pst_filename, output_dir):
    readpst_command = [
        "readpst",
        "-e",
        "-o", output_dir,
        pst_filename,
    ]

    try:
        subprocess.run(readpst_command, check=True)
        print("Email extraction completed successfully.")
    except subprocess.CalledProcessError as e:
        print(f"Error: {e}")

def find_email_files(directory, search_pattern):
    email_files = []
    
    for root, _, files in os.walk(directory):
        for file in files:
            if fnmatch.fnmatch(file, search_pattern):
                email_files.append(os.path.join(root, file))
    
    return email_files

def main():
    output_dir = os.path.dirname(os.path.abspath(__file__))
    search_pattern = "*.eml"

    pst_filename = "pass.pst"

    extract_emails_from_pst(pst_filename, output_dir)

    email_files = find_email_files(output_dir, search_pattern)
    extracted_data = []

    for email_file_path in email_files:
        with open(email_file_path, "rb") as email_file:
            email_message = email.message_from_binary_file(email_file)
            email_info = extract_email_info(email_message, email_file_path)
            extracted_data.append(email_info)
                
    output_json_file = os.path.join(output_dir, "output.json")
    
    metadata = {
        "extraction_date": str(datetime.now()),
        "extracted_files_count": len(extracted_data),
    }
    
    full_output = {
        "metadata": metadata,
        "extracted_data": extracted_data,
    }
    
    with open(output_json_file, "w") as json_file:
        json.dump(full_output, json_file, indent=4)
    
    print("Data saved to", output_json_file)


    if os.path.exists(output_json_file):
        with open(output_json_file, "r") as json_file:
            json_data = json.load(json_file)
            if "extracted_data" in json_data:
                first_email_date = json_data["extracted_data"][0]["Date"]
                first_email_date_obj = datetime.strptime(first_email_date, "%Y-%m-%d %H:%M:%S")
                year = first_email_date_obj.year
                month_number = first_email_date_obj.month
                month_name = calendar.month_name[month_number]
                new_output_json_file = os.path.join(output_dir, f"{year:04d}-{month_name}-{month_number:02d}_output.json")

                # Check if any other month is different
                any_other_month_different = any(
                    datetime.strptime(email_info["Date"], "%Y-%m-%d %H:%M:%S").month != month_number
                    for email_info in json_data["extracted_data"][1:]
                )

                if any_other_month_different:
                    print("Some of the months are different. Keeping the name as output.json")

                elif os.path.exists(new_output_json_file):
                    print(f"{new_output_json_file} already exists. Keeping the name as output.json")

                else:
                    os.rename(output_json_file, new_output_json_file)
                    print(f"Renamed output.json to {new_output_json_file}")

if __name__ == "__main__":
    main()
