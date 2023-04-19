# Get the window handles for all open windows
$handles = $driver.WindowHandles

# Switch to the last window handle (the child window)
$driver.SwitchTo().Window($handles[$handles.Count - 1])

# Get the session ID for the child window
$sessionId = $driver.SessionId

# Navigate to a URL in the child window
$driver.Navigate().GoToUrl("https://www.bing.com")

# Close the child window
$driver.Close()

# Switch back to the parent window
$driver.SwitchTo().Window($handles[0])

# Close the parent window
$driver.Close()
