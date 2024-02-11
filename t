import win32com.client

def export_rule_settings():
    # Create Outlook Application object
    outlook_app = win32com.client.Dispatch("Outlook.Application")
    
    # Get Rules object
    rules = outlook_app.Session.DefaultStore.GetRules()
    
    # Export rules to XML
    rules_xml = rules.SaveAsXML()
    
    # Write rules XML to a file
    with open("outlook_rule_settings.xml", "w") as output_file:
        output_file.write(rules_xml)
    
    print("Exported rule settings to outlook_rule_settings.xml")

if __name__ == "__main__":
    export_rule_settings()
