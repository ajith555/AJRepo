import win32com.client

# Create an instance of the Outlook application
outlook = win32com.client.Dispatch("Outlook.Application")

# Access the Rules collection
rules = outlook.Session.DefaultStore.GetRules()

# Iterate through the rules and print their names
for rule in rules:
    print("Rule Name:", rule.Name)
    print("Rule Enabled:", rule.Enabled)
    print("Rule Execution Order:", rule.ExecutionOrder)
    print("Rule Is Local:", rule.IsLocal)
    print("Rule Is Account Wide:", rule.IsAccountWide)
    
    # Print conditions
    print("Conditions:")
    for condition in rule.Conditions:
        print("- Condition:", condition.Text)

    # Print actions
    print("Actions:")
    for action in rule.Actions:
        print("- Action:", action.Name)

    print()
