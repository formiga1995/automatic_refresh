var = 0
do while var <> 2 
	var = msgbox("Atualizar as primeiras 18 abas?", vbYesNoCancel,"refreshtabs")
	if var = 6 then call refresh()
loop

function refresh()    
	ida = 0
	volta = 0
	Set object = Wscript.CreateObject("Wscript.Shell")
	wscript.sleep(3000)

	'object.appactivate "chrome.exe"
	
	do while ida < 18

		object.SendKeys "{F5}"
		wscript.sleep(125)
		object.SendKeys "^{TAB}"
	
		ida = ida+1
	loop

	do while volta < 18

		object.SendKeys "^+{TAB}"
		wscript.sleep(125)
	
		volta= volta+1
	loop
end function
