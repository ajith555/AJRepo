import win32com.client

def export_rules_to_text():
    try:
        # Create an instance of the Outlook application
        outlook = win32com.client.Dispatch("Outlook.Application")

        # Access the Rules collection
        rules = outlook.Session.DefaultStore.GetRules()

        # Open a text file to write the rule settings
        with open("outlook_rule_settings.txt", "w") as file:
            # Loop through each rule and write the name, conditions, and actions to the file
            for rule in rules:
                file.write(f"Rule Name: {rule.Name}\n")
                file.write(f"Rule Enabled: {rule.Enabled}\n")
                file.write(f"Rule Execution Order: {rule.ExecutionOrder}\n")
                
                # Check if IsLocal property exists
                if hasattr(rule, 'IsLocal'):
                    file.write(f"Rule Is Local: {rule.IsLocal}\n")
                else:
                    file.write("Rule Is Local: Not available\n")
                
                file.write(f"Rule Is Account Wide: {rule.IsAccountWide}\n")

                # Write conditions to the file
                file.write("Conditions:\n")
                for condition in rule.Conditions:
                    file.write(f"- Condition: {condition.Text}\n")

                # Write actions to the file
                file.write("Actions:\n")
                for action in rule.Actions:
                    file.write(f"- Action: {action.Name}\n")

                file.write("\n")

        print("Exported rule settings to outlook_rule_settings.txt")
    except Exception as e:
        print("Error:", e)

if __name__ == "__main__":
    export_rules_to_text()
