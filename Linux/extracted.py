import os                           # For the creating the files and the file paths
import json                         # Creatung the output.json file and reading 
from datetime import datetime       # interpreting the creation and modifcation time of files
import hashlib                      # Hashing the nessesary files 
import fnmatch                      # Matching wildcard e.g. "*.conf"
import email                        # Decoding the emails to get important information
from extract_msg import Message     # Reading the PST file
import subprocess                   # To help with error handeling
import calendar                     # Converting the date and time for the file names
import shutil                       # Moving the files

# Creating the base template for all of the logstash config files for the next steps.
# path => "{json_path}"  This is for the file path to be created later on in the script and can be replicated.
def generate_logstash_config(fields, json_path):
    config = f'''
input {{
  file {{
    path => "{json_path}"
    start_position => "beginning"
    sincedb_path => "/dev/null"
  }}
}}

filter {{
  json {{
    source => "email"
  }}

  mutate {{
    add_field => {{
'''
    for field in fields:
        config += f'      "{field}" => "%{{[extracted_data][0]["{field}"]}}"\n'
    
    config += '''    }
  }
}}

output {{
  elasticsearch {{
    hosts => ["localhost:9200"]
    index => "email_data"
  }}
  stdout {{}}
}}
'''
    return config

# This is the hashing function for all aspects of the script NOT THE PST HASH
def calculate_hash(data):
    hash_object = hashlib.sha256()

    if isinstance(data, bytes):
        hash_object.update(data)
    elif isinstance(data, str):  # Assuming data is a file path
        with open(data, 'rb') as f:
            while chunk := f.read(8192):
                hash_object.update(chunk)
    else:
        raise ValueError("Unsupported data type for hash calculation") # Error handling

    return hash_object.hexdigest()

def extract_msg_attachment_info(msg_attachment):
    msg = Message(msg_attachment)
    subject = msg.subject

    attachment_data = msg_attachment
    hash_value = calculate_hash(attachment_data)

def extract_email_info(email_message, file_path):
    sender = email_message.get("From", "Unknown Sender")
    recipient = email_message.get("To", "Unknown Recipient")
    subject = email_message.get("Subject", "No Subject")
    date_str = email_message.get("Date", "")
    Message_Id = email_message.get("Message-Id", "")

    with open(file_path, 'rb') as f:
        email_data = f.read()
    hash_value = calculate_hash(email_data)


    date_obj = email.utils.parsedate_to_datetime(date_str)
    formatted_date = date_obj.strftime("%Y-%m-%d %H:%M:%S")

    hash_value = calculate_hash(file_path)
    file_name = os.path.basename(file_path)

    # Extract attachment information
    # Need to add for attachments attachment if the fisrt one was a .msg
    attachments = []
    for part in email_message.walk():
        if part.get_content_maintype() == "application" and part.get_content_disposition() == "attachment":
            attachment_name = part.get_filename()
            attachment_data = part.get_payload(decode=True)

            if attachment_name.lower().endswith(".msg"): # Need to check syntax for doing more into this .msg file. 
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

def extract_emails_from_pst(pst_filename, output_dir): # This is the basic linux command to run the script 
    readpst_command = [
        "readpst",
        "-e",
        "-o", output_dir,
        pst_filename,
    ]

    try:
        subprocess.run(readpst_command, check=True) # Main part of the script running.
        print("Email extraction completed successfully.") # Print to check output was completed.
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
    search_pattern = "*.eml" # pst opens the files as a .eml 
    with open("upload/uploaded_log.txt", "r") as f:
        entries = f.readlines()
        print(entries)
    pst_filename = os.path.join("upload", entries[0].strip())
 # Extract the first entry and remove leading/trailing whitespace
    # See link https://machinelearningmastery.com/command-line-arguments-for-your-python-script/
    

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

    # Find the JSON file in the directory where the script is located
    json_files = [filename for filename in os.listdir('.') if filename.endswith('.json')]
    if len(json_files) == 0:
        print("No JSON files found.")
        return

    json_file = json_files[0]  # Assuming there's only one JSON file
    print(f"Using {json_file} as the JSON file.")
    fields_to_use = ["Sender", "Recipient", "Subject", "Date", "Hash", "Message Id", "File Name"]
    # Read the JSON file
    with open(json_file, "r") as json_file:
        json_data = json.load(json_file)
        extracted_data = json_data["extracted_data"]
        for email_info in extracted_data:
            json_path = email_info["File Name"]  # Use the appropriate field here
            base_filename = os.path.splitext(os.path.basename(json_path))[0]
            config_content = generate_logstash_config(fields_to_use, json_path)
            
            # Write the generated Logstash configuration to a file
            config_filename = f"{base_filename}_config.conf"
            with open(config_filename, "w") as config_file:
                config_file.write(config_content)

            print(f"Generated configuration for {json_path} as {config_filename}")
    for email_info in extracted_data:
        json_path = email_info["File Name"]  # Use the appropriate field here
        base_filename = os.path.splitext(os.path.basename(json_path))[0]
        config_content = generate_logstash_config(fields_to_use, json_path)
        
        # Write the generated Logstash configuration to a file
        config_filename = f"{base_filename}_config.conf"
        with open(config_filename, "w") as config_file:
            config_file.write(config_content)

        print(f"Generated configuration for {json_path} as {config_filename}")

        # Move JSON files with .conf extension to the 'config' subfolder
        if config_filename.endswith(".conf"):
            config_subfolder = os.path.join(output_dir, "config")
            os.makedirs(config_subfolder, exist_ok=True)
            config_dest_path = os.path.join(config_subfolder, config_filename)
            shutil.move(config_filename, config_dest_path)
            print(f"Moved {config_filename} to {config_dest_path}")

if __name__ == "__main__":
    main()
