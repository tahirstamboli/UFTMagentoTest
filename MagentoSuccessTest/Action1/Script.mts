﻿﻿set oShell = CreateObject("WScript.Shell")
APP_URL = oShell.ExpandEnvironmentStrings("%TEST_URL%")
Call Validate_home_page_magento(APP_URL)
'Call Validate_Add_to_Cart


Function Validate_home_page_magento(APP_URL)
Set oIE = CreateObject("InternetExplorer.Application")
oIE.Visible = True
oIE.Navigate(APP_URL)

'Maximize the Browser Window
Window("hwnd:=" & oIE.HWND).Maximize

' Test Case to check if home page is Displayed
If Browser("Title:=Home page*").Page("title:=Home Page*").Exist(1) Then
'Print a Pass message 
Print "User is on Home Page"
Else
Print "Error !!!"
End If

Browser("Home page").Page("Home page").WebEdit("q").Set "Samsung" @@ hightlight id_;_Browser("Home page").Page("Home page").WebEdit("q")_;_script infofile_;_ZIP::ssf88.xml_;_
Browser("Home page").Page("Home page").WebButton("Search").Click @@ hightlight id_;_Browser("Home page").Page("Home page").WebButton("Search")_;_script infofile_;_ZIP::ssf89.xml_;_
wait 10

End Function @@ hightlight id_;_Browser("Home page").Page("Home page").WebButton("Search")_;_script infofile_;_ZIP::ssf89.xml_;_

Function Validate_Add_to_Cart
Browser("Home page").Page("Search results for: 'samsung'").WebButton("Add to Cart").Click @@ hightlight id_;_Browser("Home page").Page("Search results for: 'samsung'").WebButton("Add to Cart")_;_script infofile_;_ZIP::ssf90.xml_;_
wait 20
Call CaptureScreenSnap_homepage()
wait 10
Browser("Home page").Quit
End Function


Function CaptureScreenSnap_homepage
Browser("Home page").CaptureBitmap "C:\Report\Magento.png", True
End Function
