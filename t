import os
import re

def export_rule_settings():
    # Get the Outlook data directory
    outlook_data_dir = os.path.join(os.getenv('APPDATA'), 'Microsoft', 'Outlook')
    
    # Look for the rules file
    rules_file_path = None
    for file_name in os.listdir(outlook_data_dir):
        if file_name.lower().startswith('outlook') and file_name.lower().endswith('.dat'):
            rules_file_path = os.path.join(outlook_data_dir, file_name)
            break
    
    if not rules_file_path:
        print("Outlook rules file not found.")
        return
    
    # Read the rules file
    with open(rules_file_path, 'rb') as file:
        rules_data = file.read().decode('utf-16-le')
    
    # Extract rule settings
    rule_settings = re.findall(r'\[RULE\].*?\[/RULE\]', rules_data, re.DOTALL)
    
    # Write rule settings to a text file
    with open("outlook_rule_settings.txt", "w") as output_file:
        for rule_setting in rule_settings:
            output_file.write(rule_setting)
            output_file.write("\n\n")
    
    print("Exported rule settings to outlook_rule_settings.txt")

if __name__ == "__main__":
    export_rule_settings()
