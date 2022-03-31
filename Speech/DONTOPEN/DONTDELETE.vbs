x=msgbox("Press 'OK' to start Speech.", 0+32, "Speech")

Dim message, sapi
message=InputBox("Type what you want to hear.","Speech")
Set sapi=CreateObject("sapi.spvoice")
sapi.Speak message