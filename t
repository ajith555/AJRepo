import os
import re

def export_rule_settings():
    # Get the Outlook data directory
    outlook_data_dir = os.path.join(os.getenv('APPDATA'), 'Microsoft', 'Outlook')
    
    # Look for files containing rule settings
    rule_files = []
    for file_name in os.listdir(outlook_data_dir):
        if file_name.lower().endswith('.dat'):
            rule_files.append(os.path.join(outlook_data_dir, file_name))
    
    if not rule_files:
        print("No Outlook rule files found.")
        return
    
    # Read rule settings from each file
    rule_settings = []
    for file_path in rule_files:
        with open(file_path, 'rb') as file:
            rules_data = file.read().decode('utf-16-le')
            rule_settings.extend(re.findall(r'\[RULE\].*?\[/RULE\]', rules_data, re.DOTALL))
    
    # Write rule settings to a text file
    with open("outlook_rule_settings.txt", "w") as output_file:
        for rule_setting in rule_settings:
            output_file.write(rule_setting)
            output_file.write("\n\n")
    
    print("Exported rule settings to outlook_rule_settings.txt")

if __name__ == "__main__":
    export_rule_settings()
