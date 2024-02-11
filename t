# Provide the path to the .rwz file
$rwzFilePath = "C:\path\to\your\rules.rwz"

# Check if the file exists
if (Test-Path $rwzFilePath) {
    # Load the Outlook rules from the .rwz file
    $outlook = New-Object -ComObject Outlook.Application
    $rules = $outlook.Session.CreateObjectFromOutlookTemplate($rwzFilePath).Rules

    # Open a file to write the rule settings
    $outfile = "outlook_rule_settings.txt"
    $stream = [System.IO.StreamWriter] $outfile

    # Loop through each rule and write the name, conditions, and actions to the file
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
        foreach ($condition in $rule.Conditions) {
            $description += "Condition: $($condition.Text)`n"
        }

        # Construct actions part of the description
        $description += "Actions:`n"
        foreach ($action in $rule.Actions) {
            $description += "Action: $($action.Text)`n"
        }

        # Write the rule description to the file
        $stream.WriteLine($description)
        $stream.WriteLine("`n")
    }

    # Close the file
    $stream.Close()

    Write-Host "Exported rule settings to $outfile"
} else {
    Write-Host "File not found at specified location: $rwzFilePath"
}
