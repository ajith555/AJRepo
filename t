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

    # Access various properties of the rule
    $description += "Rule Enabled: $($rule.Enabled)`n"
    $description += "Rule Execution Order: $($rule.ExecutionOrder)`n"
    $description += "Rule Is Local: $($rule.IsLocal)`n"
    $description += "Rule Is Account Wide: $($rule.IsAccountWide)`n"

    # Construct conditions part of the description
    $description += "Conditions:`n"
    if ($rule.Conditions -eq $null) {
        $description += "No conditions defined`n"
    } else {
        foreach ($condition in $rule.Conditions | Where-Object { $_.Enabled }) {
            $description += "Condition: $($condition.Text)`n"
        }
    }

    # Construct actions part of the description
    $description += "Actions:`n"
    if ($rule.Actions -eq $null) {
        $description += "No actions defined`n"
    } else {
        foreach ($action in $rule.Actions | Where-Object { $_.Enabled }) {
            $description += "Action: $($action.Name)`n"
        }
    }

    # Write the rule description to the file
    $stream.WriteLine($description)
    $stream.WriteLine("`n")
}

# Close the file
$stream.Close()

Write-Host "Exported rule settings to $outfile"
