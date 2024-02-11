import win32com.client

def export_rules():
    # Create Outlook Application object
    outlook_app = win32com.client.Dispatch("Outlook.Application")
    
    # Get Rules object
    rules = outlook_app.Session.DefaultStore.GetRules()
    
    # Open a file to write the rules
    with open("outlook_rules.txt", "w") as file:
        for rule in rules:
            # Print available attributes
            print(dir(rule))
            
            # Write rule name and description to the file
            file.write(f"Rule Name: {rule.Name}\n")
            # Use a try-except block to handle the case where Description attribute doesn't exist
            try:
                file.write(f"Description: {rule.Description}\n\n")
            except AttributeError:
                file.write("Description: Not available\n\n")
    
    print("Exported rules to outlook_rules.txt")

if __name__ == "__main__":
    export_rules()
