# Install Selenium module if not already installed
if (-not (Get-Module -Name Selenium)) {
    Install-Module -Name Selenium
}

# Import Selenium module
Import-Module Selenium

# Path to chromedriver executable
$ChromeDriverPath = "C:\path\to\chromedriver.exe"

# Start a Selenium Chrome driver session
$ChromeOptions = New-Object OpenQA.Selenium.Chrome.ChromeOptions
$ChromeOptions.AddArgument("--headless") # Optional: Run headless to not show browser window
$Driver = New-Object OpenQA.Selenium.Chrome.ChromeDriver($ChromeDriverPath, $ChromeOptions)

# Navigate to the website
$Driver.Url = "https://ubs.invisionapp.com"

# Add delay to allow time for page to load
Start-Sleep -Seconds 5

# Perform authentication using Selenium (add your own logic here)
# For example, find username and password fields and input credentials, then submit

# Wait for authentication to complete (add your own logic here)
# For example, wait for certain element to appear indicating successful login

# Fetch all URLs matching the pattern
$URLs = $Driver.FindElementsByXPath("//a[contains(@href,'/overview/') and contains(@href,'/exports/initiate/pdf?method=overviewMenu&sortBy=1&sortOrder=1&ViewLayout=2')]")
$URLList = @()
foreach ($URL in $URLs) {
    $URLList += $URL.GetAttribute("href")
}

# Close the Selenium Chrome driver session
$Driver.Quit()

# Output the list of URLs
$URLList
