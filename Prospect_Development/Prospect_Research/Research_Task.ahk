; Context: You are entering a Task into Advance.
; Initiate the Script with WIN + t
; Advance should be open to the record you want to enter the task for
#A::
Reload
Return

#IfWinActive Production Database
#t::
CHAR_LIM := 255
Gui, New
Gui, Font, s28 bold
Gui, Add, Text, w300 h200, Prospect Research`n`nEnter Task:
Gui, Font, s10 normal
Gui, Add, Text,, Type
Gui, Add, DropDownList, vType, Identification||CapacityUpgrade
Gui, Add, Text,, Justification
Gui, Add, Edit, r3 w275 vTaskJustification -WantReturn
Gui, Add, Button, default, OK
Gui, Show
Return

GuiClose:
ButtonOK:
    Gui, Submit

Sleep, 100

; Open GOTO Window
Send, {F5}

; Enter code for Tasks
Send, ENTT
Send, {ENTER}


; Open new Task
Send, {F6}

if (Type = "Identification")
{
    SendInput, ID
    Sleep, 100
} else
{
    SendInput, RA
    Sleep, 100
}

Send, {TAB}

Send, 1
SendInput, {TAB 4}

Send, 600768
SendInput, {TAB 6}

FormatTime, CurrentDateTime,, MM/dd/yy
SendInput %CurrentDateTime%
Sleep, 100

SendInput, {TAB 5}
Sleep, 100
Send, %TaskJustification% 
Send, {TAB}

Send, RE
Sleep, 50
SendInput, {TAB 2}

Send, 649757
return