Call Validate_Add_to_Cart

Function Validate_Add_to_Cart
'Updated by test maintenance run
Browser("Home page").Page("Search results for: 'samsung'").WebButton("Add to Cart").Click
'Browser("Home page").Page("Search results for: 'samsung'").WebElement("Add to Cart").Click
wait 20
'Verify user is on add to cart page
If Browser("Title:=Home page*").Page("title:=Home Page*").Exist(1) Then
Print "User is on Home Page"
'Capture screenshot for Home Page 
Call CaptureScreen_homepage
Else
Print "Error !!!"
End If
Call CaptureScreen_cart()
wait 10
Browser("Home page").Quit
End Function

Function CaptureScreen_cart
Browser("Home page").CaptureBitmap "C:\testnow\code\Report\Magento_Cart.png", True
End Function
