# Install Selenium module if not already installed
if (-not (Get-Module -Name Selenium.WebDriver)) {
    Install-Module -Name Selenium.WebDriver
}

# Import Selenium module
Import-Module Selenium.WebDriver

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

# Find username and password fields and input credentials
$UsernameField = $Driver.FindElementById("username") # Replace "username" with actual ID of username field
$PasswordField = $Driver.FindElementById("password") # Replace "password" with actual ID of password field

$UsernameField.SendKeys("your_username") # Replace "your_username" with actual username
$PasswordField.SendKeys("your_password") # Replace "your_password" with actual password

# Find login button and click
$LoginButton = $Driver.FindElementById("loginButton") # Replace "loginButton" with actual ID of login button
$LoginButton.Click()

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
