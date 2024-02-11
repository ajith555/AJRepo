import win32com.client

def get_rule_description(rule):
    description = f"Apply this rule after message arrives"
    
    # Construct conditions part of the description
    conditions = getattr(rule, "Conditions", None)
    if conditions:
        for condition in conditions:
            if condition.Enabled:
                if condition.ConditionType == 0:  # Subject
                    description += f" with '{condition.Text}' in the subject"
                # Add more conditions as needed
            
    # Construct actions part of the description
    actions = getattr(rule, "Actions", None)
    if actions:
        for action in actions:
            if action.Enabled:
                if action.ActionType == 0:  # Move to folder
                    description += f" move it to '{action.Folder.Name}' folder"
                # Add more actions as needed
                
    # Add any additional rule settings
    if rule.StopProcessingRule:
        description += " and stop processing more rules"
    
    return description

def export_rule_settings():
    # Create Outlook Application object
    outlook_app = win32com.client.Dispatch("Outlook.Application")
    
    # Get Rules object
    rules = outlook_app.Session.DefaultStore.GetRules()
    
    # Open a file to write the rule settings
    with open("outlook_rule_settings.txt", "w") as output_file:
        for rule in rules:
            # Write rule name and constructed description to the file
            output_file.write(f"Rule Name: {rule.Name}\n")
            output_file.write(f"Rules Description: {get_rule_description(rule)}\n\n")
    
    print("Exported rule settings to outlook_rule_settings.txt")

if __name__ == "__main__":
    export_rule_settings()
