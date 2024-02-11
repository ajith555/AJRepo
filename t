import win32com.client

def export_rule_settings():
    # Create Outlook Application object
    outlook_app = win32com.client.Dispatch("Outlook.Application")
    
    # Get Rules object
    rules = outlook_app.Session.DefaultStore.GetRules()
    
    # Open a file to write the rule settings
    with open("outlook_rule_settings.txt", "w") as output_file:
        for rule in rules:
            # Write rule name and description to the file
            output_file.write(f"Rule Name: {rule.Name}\n")
            output_file.write(f"Rules Description: {rule.Text}\n\n")
    
    print("Exported rule settings to outlook_rule_settings.txt")

if __name__ == "__main__":
    export_rule_settings()
