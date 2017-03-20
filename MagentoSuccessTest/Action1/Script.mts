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

' Test Case to check if home page is Displayed
If Browser("Title:=Home page*").Page("title:=Home Page*").Exist(1) Then
'Print a Pass message 
Print "User is on Home Page"
'Capture screenshot for Home Page 
Call CaptureScreen_homepage
Else
Print "Error !!!"
End If

'wait 10

Browser("Home page").Page("Home page").WebEdit("q").Set "Samsung" @@ hightlight id_;_Browser("Home page").Page("Home page").WebEdit("q")_;_script infofile_;_ZIP::ssf88.xml_;_
Browser("Home page").Page("Home page").WebButton("Search").Click @@ hightlight id_;_Browser("Home page").Page("Home page").WebButton("Search")_;_script infofile_;_ZIP::ssf89.xml_;_
Wait 5


Browser("Internet Explorer Enhanced").Dialog("Internet Explorer").WinButton("Yes").Click @@ hightlight id_;_6228634_;_script infofile_;_ZIP::ssf99.xml_;_

wait 5

End Function 

Function CaptureScreen_homepage
Browser("Home page").CaptureBitmap "C:\testnow\code\Report\Magento_home_page.png", True
End Function
