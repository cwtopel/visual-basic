Option Explicit
Dim name

'InputBox "message", "title", "input field", xposition, yposition

name = InputBox("What is your name?", "Information:", "Names goes here.", 15000, 10000)
'msgbox name, vbOKOnly, "This is your name?"

if name = "Jeremy" or name = "Chris" then
	msgbox "Hello!"
elseif name <> "Jeremy" then
	msgbox "INTRUDER"
end if