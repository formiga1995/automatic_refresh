ida = 0
volta = 0
Set object = Wscript.CreateObject("Wscript.Shell")
wscript.sleep(3000)

'object.appactivate "chrome.exe"

do while ida < 15

	object.SendKeys "{F5}"
	wscript.sleep(125)
	object.SendKeys "^{TAB}"
	
	ida = ida+1
loop

do while volta < 15

	object.SendKeys "^+{TAB}"
	wscript.sleep(125)
	
	volta= volta+1
loop

