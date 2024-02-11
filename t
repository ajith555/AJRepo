import win32com.client

def construct_rule_description(rule):
    description = f"Apply this rule after message arrives"
    
    # Construct conditions part of the description
    conditions = getattr(rule, "Conditions", None)
    if conditions:
        for condition in conditions:
            # Example: Add condition for subject
            if condition.Enabled:
                if condition.ConditionType == 0:  # Subject
                    description += f" with '{condition.Text}' in the subject"
                # Add more conditions as needed
            
    # Construct actions part of the description
    actions = getattr(rule, "Actions", None)
    if actions:
        for action in actions:
            # Example: Add action for moving to a folder
            if action.Enabled:
                if action.ActionType == 0:  # Move to folder
                    description += f" move it to '{action.Folder.Name}' folder"
                # Add more actions as needed
                
    # Add any additional rule settings
    if rule.StopProcessingRule:
        description += " and stop processing more rules"
    
    return description

def export_rules():
    # Create Outlook Application object
    outlook_app = win32com.client.Dispatch("Outlook.Application")
    
    # Get Rules object
    rules = outlook_app.Session.DefaultStore.GetRules()
    
    # Open a file to write the rules
    with open("outlook_rules.txt", "w") as file:
        for rule in rules:
            # Write rule name and constructed description to the file
            file.write(f"Rule Name: {rule.Name}\n")
            file.write(f"Rules Description: {construct_rule_description(rule)}\n\n")
    
    print("Exported rules to outlook_rules.txt")

if __name__ == "__main__":
    export_rules()
