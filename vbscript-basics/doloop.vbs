Option Explicit
Dim a, pass
a = 0

do until a = 5 'do until a>4 OR do while a<6
	a = a+1
	msgbox a
loop

do
pass = inputbox("Password")
if pass = "wired" then
exit do
elseif pass = "" then
	msgbox "Don't leave the field blank."
elseif pass <> "wired" then
	msgbox "Incorrect!", vbCritical
end if
loop

msgbox "Correct!"