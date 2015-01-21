Option Explicit
Dim a

'a = msgbox ("Pandora Radio", vbAbortRetryIgnore+vbExclamation+vbDefaultButton2+vbSystemModal, "Gadget:")
'if a = vbAbort then
'	msgbox "Quit", vbCritical
'elseif a = vbRetry then
'	msgbox "Quit2", vbQuestion
'end if

a = msgbox("Guess a button.", vbAbortRetryIgnore)
if a = vbRetry or a = vbAbort then
	msgbox "Correct!"
else
	msgbox "Wrong!!"
end if