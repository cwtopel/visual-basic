a = msgbox ("Pandora Radio", vbAbortRetryIgnore+vbExclamation+vbDefaultButton2+vbSystemModal, "Gadget:")
if a = vbAbort then msgbox "Quit", vbCritical
if a = vbRetry then msgbox "Quit2", vbQuestion