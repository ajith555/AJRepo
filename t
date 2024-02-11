import win32com.client

def export_rules():
    # Create Outlook Application object
    outlook_app = win32com.client.Dispatch("Outlook.Application")
    
    # Get Rules object
    rules = outlook_app.Session.DefaultStore.GetRules()
    
    # Open a file to write the rules
    with open("outlook_rules.txt", "w") as file:
        for rule in rules:
            # Write rule name to the file
            file.write(f"Rule Name: {rule.Name}\n")
            
            # Print Conditions
            file.write("\nConditions:\n")
            conditions = getattr(rule, "Conditions", None)
            if conditions:
                file.write(f"{conditions}\n")
            else:
                file.write("No conditions found.\n")
            
            # Print Actions
            file.write("\nActions:\n")
            actions = getattr(rule, "Actions", None)
            if actions:
                for action in actions:
                    file.write(f"{action}\n")
            else:
                file.write("No actions found.\n")
            
            file.write("\n\n")
    
    print("Exported rules to outlook_rules.txt")

if __name__ == "__main__":
    export_rules()
