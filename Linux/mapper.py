#This will only work with the running extracted.py is completed and successful 
#
import os
from difflib import SequenceMatcher

def calculate_similarity(reference_config, config):
    matcher = SequenceMatcher(None, reference_config, config)
    similarity_ratio = matcher.ratio() * 100
    return similarity_ratio

def main():
    output_directory = "config/"  # need to make a var based on the extracted.py file 

    # Reference configuration template
    reference_config_template = """
input {
  file {
    path => [%s]
    start_position => "beginning"
    sincedb_path => "/dev/null"
  }
}

filter {
  json {
    source => "email"
  }

  mutate {
    add_field => {
      "Sender" => "%%{[extracted_data][0]['Sender']}"
      "Recipient" => "%%{[extracted_data][0]['Recipient']}"
      "Subject" => "%%{[extracted_data][0]['Subject']}"
      "Date" => "%%{[extracted_data][0]['Date']}"
      "Hash" => "%%{[extracted_data][0]['Hash']}"
      "Message Id" => "%%{[extracted_data][0]['Message Id']}"
      "File Name" => "%%{[extracted_data][0]['File Name']}"
    }
  }
}}

output {
  elasticsearch {
    hosts => ["localhost:9200"]
    index => "email_data"
  }
  stdout {}
}
"""

    config_files = []
    matching_paths = []

    for subdir, _, files in os.walk("config"): # need to make a var 
        for file in files:
            if file.endswith('.conf'):
                config_file_path = os.path.join(subdir, file)
                with open(config_file_path, "r") as config_file:
                    config_data = config_file.read()  # Read the entire contents of the file

                similarity = calculate_similarity(reference_config_template, config_data)
                print(f"Matching percentage for {file}: {similarity:.2f}%")
                
                if similarity >= 58: # Keep at this value
                    matching_paths.append(f'"{os.path.join("Outlook Data File", "Inbox", file)}"') # need to make a var 

    combined_config = reference_config_template % (", ".join(matching_paths))

    with open(os.path.join(output_directory, "combined_config.conf"), "w") as combined_file:
        combined_file.write(combined_config)

if __name__ == "__main__":
    main()
