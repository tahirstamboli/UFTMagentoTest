Call Validate_Add_to_Cart

Function Validate_Add_to_Cart
'Updated by test maintenance run
Browser("Home page").Page("Search results for: 'samsung'").WebButton("Add to Cart").Click
'Browser("Home page").Page("Search results for: 'samsung'").WebElement("Add to Cart").Click
wait 20
Call CaptureScreenSnap_homepage()
wait 10
Browser("Home page").Quit
End Function

Function CaptureScreenSnap_homepage
Browser("Home page").CaptureBitmap "C:\testnow\code\Report\Magento_add_to_cart.png", True
End Function
