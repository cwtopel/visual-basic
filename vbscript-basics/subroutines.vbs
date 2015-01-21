Option Explicit
Dim x

sub Greeting(title, message)
msgbox "hello", message, title
end sub

sub Finish
msgbox "Goodbye"
end sub

call Greeting("Welcome!", 20)
call Finish