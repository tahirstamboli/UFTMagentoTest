set oShell = CreateObject("WScript.Shell")
APP_URL = oShell.ExpandEnvironmentStrings("%TEST_URL%")
'BROWSE = oShell.Environment("User").Item("BROWSER")
BROWSE = oShell.ExpandEnvironmentStrings("%BROWSER%")
Print "The Application URL is : " & APP_URL
Print "The Browser Passed is: " & BROWSE
Call Validate_home_page_magento(APP_URL,BROWSE)



Function Validate_home_page_magento(APP_URL,Search_engine)
set objShell = CreateObject("Shell.Application")

If Search_engine="Chrome" Then
	   objShell.ShellExecute "chrome.exe", APP_URL, "", "", 1
		Print "Chrome Browser Selected"
   
   ElseIF Search_engine="Firefox" Then
		objShell.ShellExecute "firefox.exe", APP_URL, "", "", 1
		Print "Firefox Browser Selected"
   
   ElseIf Search_engine="IE" Then
	    Set oIE = CreateObject("InternetExplorer.Application")
	    oIE.Visible = True
	    oIE.Navigate(APP_URL)
	    Print "IE Browser Selected"
	    Window("hwnd:=" & oIE.HWND).Maximize
End If

wait 40
  
End Function 
