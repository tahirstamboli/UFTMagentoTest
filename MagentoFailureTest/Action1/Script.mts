set oShell = CreateObject("WScript.Shell")

TEST_URL = oShell.ExpandEnvironmentStrings("%TEST_URL%")
'Browser =  oShell.ExpandEnvironmentStrings("%Browser%")
'TEST_URL = ("http://54.89.102.77")
Call Magento_Search_Validation(TEST_URL)

 @@ hightlight id_;_Browser("Home page").Page("Checkout").WebButton("Place Order")_;_script infofile_;_ZIP::ssf50.xml_;_
Function Magento_Search_Validation(TEST_URL)
Set oIE = CreateObject("InternetExplorer.Application")
oIE.Visible = True
oIE.Navigate(TEST_URL)

Browser("Home page").Page("Home page").WebEdit("q").Set "Samsung" @@ hightlight id_;_Browser("Home page").Page("Home page").WebEdit("q")_;_script infofile_;_ZIP::ssf88.xml_;_
Browser("Home page").Page("Home page").WebButton("Search").Click @@ hightlight id_;_Browser("Home page").Page("Home page").WebButton("Search")_;_script infofile_;_ZIP::ssf89.xml_;_
Browser("Home page").Page("Search results for: 'samsung'").WebButton("Add to Cart").Click @@ hightlight id_;_Browser("Home page").Page("Search results for: 'samsung'").WebButton("Add to Cart")_;_script infofile_;_ZIP::ssf90.xml_;_
Call CaptureScreenSnap()
Browser("Home page").Page("Shopping Cart").WebButton("Proceed to Checkout_2").Click @@ hightlight id_;_Browser("Home page").Page("Shopping Cart").WebButton("Proceed to Checkout 2")_;_script infofile_;_ZIP::ssf91.xml_;_
Browser("Home page").Page("Checkout").WebRadioGroup("checkout_method").Select "register" @@ hightlight id_;_Browser("Home page").Page("Checkout").WebRadioGroup("checkout method")_;_script infofile_;_ZIP::ssf92.xml_;_
Browser("Home page").Page("Checkout").WebButton("Continue").Click @@ hightlight id_;_Browser("Home page").Page("Checkout").WebButton("Continue")_;_script infofile_;_ZIP::ssf93.xml_;_
Browser("Home page").Page("Checkout").WebEdit("billing[firstname]").Set "Paras" @@ hightlight id_;_Browser("Home page").Page("Checkout").WebEdit("billing[firstname]")_;_script infofile_;_ZIP::ssf94.xml_;_
Browser("Home page").Page("Checkout").WebEdit("billing[lastname]").Set "Mishra" @@ hightlight id_;_Browser("Home page").Page("Checkout").WebEdit("billing[lastname]")_;_script infofile_;_ZIP::ssf95.xml_;_
Browser("Home page").Page("Checkout").WebEdit("billing[company]").Set "ReanCloud" @@ hightlight id_;_Browser("Home page").Page("Checkout").WebEdit("billing[company]")_;_script infofile_;_ZIP::ssf96.xml_;_
Browser("Home page").Page("Checkout").WebEdit("billing[email]").Set "paras.mishra@reancloud.com" @@ hightlight id_;_Browser("Home page").Page("Checkout").WebEdit("billing[email]")_;_script infofile_;_ZIP::ssf97.xml_;_
Browser("Home page").Page("Checkout").WebEdit("billing[street][]").Set "Hinjewadi" @@ hightlight id_;_Browser("Home page").Page("Checkout").WebEdit("billing[street][]")_;_script infofile_;_ZIP::ssf98.xml_;_
Browser("Home page").Page("Checkout").WebEdit("billing[city]").Set "Pune" @@ hightlight id_;_Browser("Home page").Page("Checkout").WebEdit("billing[city]")_;_script infofile_;_ZIP::ssf99.xml_;_
Browser("Home page").Page("Checkout").WebEdit("billing[postcode]").Set "411057" @@ hightlight id_;_Browser("Home page").Page("Checkout").WebEdit("billing[postcode]")_;_script infofile_;_ZIP::ssf100.xml_;_
Browser("Home page").Page("Checkout").WebList("billing[country_id]").Select "Iceland" @@ hightlight id_;_Browser("Home page").Page("Checkout").WebList("billing[country id]")_;_script infofile_;_ZIP::ssf101.xml_;_
Browser("Home page").Page("Checkout").WebList("billing[country_id]").Select "India" @@ hightlight id_;_Browser("Home page").Page("Checkout").WebList("billing[country id]")_;_script infofile_;_ZIP::ssf102.xml_;_
Browser("Home page").Page("Checkout").WebEdit("billing[telephone]").Set "9011960549" @@ hightlight id_;_Browser("Home page").Page("Checkout").WebEdit("billing[telephone]")_;_script infofile_;_ZIP::ssf103.xml_;_
Browser("Home page").Page("Checkout").WebEdit("billing[customer_password]").SetSecure "58808881dd1f2e0bbce1ca036d2ad3b46e359d30cf10889d" @@ hightlight id_;_Browser("Home page").Page("Checkout").WebEdit("billing[customer password]")_;_script infofile_;_ZIP::ssf104.xml_;_
Browser("Home page").Page("Checkout").WebEdit("billing[confirm_password]").SetSecure "5880888dfbd204a9df3c9cda6705676e25eeab759d078c20" @@ hightlight id_;_Browser("Home page").Page("Checkout").WebEdit("billing[confirm password]")_;_script infofile_;_ZIP::ssf105.xml_;_
Browser("Home page").Page("Checkout").WebButton("Continue_2").Click @@ hightlight id_;_Browser("Home page").Page("Checkout").WebElement("checkout-step-shipping method")_;_script infofile_;_ZIP::ssf107.xml_;_
Browser("Home page").Quit
End Function
Function CaptureScreenSnap
Browser("Home page").CaptureBitmap "C:\Report\Magento.png", True
End Function

