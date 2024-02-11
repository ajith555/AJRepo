# Create a new Outlook application object
$outlook = New-Object -ComObject Outlook.Application

# Get the rules from the default store
$rules = $outlook.Session.DefaultStore.GetRules()

# Open a file to write the rule settings
$outfile = "outlook_rule_settings.txt"
$stream = [System.IO.StreamWriter] $outfile

# Loop through each rule and write the name and description to the file
foreach ($rule in $rules) {
    $description = "Rule Name: $($rule.Name)`n"
    $description += "Rule Description:`n"
    
    # Construct rule conditions
    $conditions = ""
    foreach ($condition in $rule.Conditions) {
        if ($condition.Enabled) {
            $conditions += "Condition: $($condition.EnabledCondition.Text)`n"
            # Add more conditions as needed
        }
    }
    
    # Construct rule actions
    $actions = ""
    foreach ($action in $rule.Actions) {
        if ($action.Enabled) {
            $actions += "Action: $($action.ActionType)`n"
            # Add more actions as needed
        }
    }
    
    # Combine conditions and actions
    $description += $conditions
    $description += $actions
    
    $stream.WriteLine($description)
}

# Close the file
$stream.Close()

Write-Host "Exported rule settings to $outfile"
