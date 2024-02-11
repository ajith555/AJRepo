import os
import re

def find_outlook_rules_file():
    # List of possible directories where the rules file might be located
    possible_directories = [
        os.path.join(os.getenv('APPDATA'), 'Microsoft', 'Outlook'),  # Default Outlook data directory
        os.path.join(os.getenv('LOCALAPPDATA'), 'Microsoft', 'Outlook'),  # LocalAppData directory
    ]
    
    # List of possible file names for the rules file
    possible_file_names = [
        'outlook.rules',  # Common Outlook rules file name
        'outlook.dat',    # Another common Outlook rules file name
    ]
    
    # Look for the rules file
    rules_file_path = None
    for directory in possible_directories:
        for file_name in possible_file_names:
            file_path = os.path.join(directory, file_name)
            if os.path.isfile(file_path):
                rules_file_path = file_path
                break
        if rules_file_path:
            break
    
    return rules_file_path

def export_rule_settings():
    # Find the Outlook rules file
    rules_file_path = find_outlook_rules_file()
    
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
