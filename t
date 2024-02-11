import win32com.client

def export_rules():
    # Create Outlook Application object
    outlook_app = win32com.client.Dispatch("Outlook.Application")
    
    # Get Rules object
    rules = outlook_app.Session.DefaultStore.GetRules()
    
    # Open a file to write the rules
    with open("outlook_rules.txt", "w") as file:
        for rule in rules:
            # Print available properties and methods
            file.write(f"Rule Name: {rule.Name}\n")
            file.write("Available Properties and Methods:\n")
            file.write("\n".join(dir(rule)))
            file.write("\n\n")
    
    print("Exported rules to outlook_rules.txt")

if __name__ == "__main__":
    export_rules()
