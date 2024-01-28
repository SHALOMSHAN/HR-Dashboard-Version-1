' Power BI Refresh and Publish Script

Dim objShell
Set objShell = WScript.CreateObject("WScript.Shell")

' Path to Power BI Desktop executable
Dim pbidePath
pbidePath = "C:\Program Files (x86)\Microsoft Power BI Desktop\bin\PBIDesktop.exe"

' Path to your Power BI Desktop file
Dim pbixFilePath
pbixFilePath = "C:\Users\shalo\Desktop\HR Dashboard Version 1\HR Dashboard.pbix"

' Sleep for a few seconds to ensure Power BI Desktop is fully loaded
WScript.Sleep 5000

' Launch Power BI Desktop and open the report
objShell.Run """" & pbidePath & """ """ & pbixFilePath & """"

' Sleep for a few seconds to allow the report to open
WScript.Sleep 5000

' Send Alt + F5 key combination to refresh the report
objShell.SendKeys "%{F5}"

' Sleep for a few seconds to allow the refresh to complete
WScript.Sleep 10000

' Send Ctrl + S key combination to save the report
objShell.SendKeys "^s"

' Sleep for a few seconds to allow the save to complete
WScript.Sleep 5000

' Send Alt + F4 key combination to close Power BI Desktop
objShell.SendKeys "%{F4}"

Set objShell = Nothing
