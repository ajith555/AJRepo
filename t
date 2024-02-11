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

    # Access the rule's property to get the rule description
    $ruleDescription = $rule.Description

    # Write the rule description to the file
    $stream.WriteLine($description)
    $stream.WriteLine($ruleDescription)
    $stream.WriteLine("`n")
}

# Close the file
$stream.Close()

Write-Host "Exported rule settings to $outfile"
