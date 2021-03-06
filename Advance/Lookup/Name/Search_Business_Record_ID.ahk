
;Context: You are search first and last names in advance
;	Your source is a spreadsheet with first and last side by side
;	the script copies last, then first and tabs over to advance 
;	any oopen window in advance is closed, and then an entity search
;	window is opened
;	Currently this script deletes two invisible characters copied from 
;	The excel file. I haven't found another way to get around it, yet
;Instructions: have advance and excel open, click advance, then excel, so they
;	are the last two programs you've accessed. select the cell of the first 
;	last name you'd like to search, and press Windows Key plus F
;	alt tab back to excel when the new search is to begin

#A::
;MsgBox A started
SetTitleMatchMode, 2

callGoToBlank(tableText, entID)
{
    SetTitleMatchMode, 2
    Send, {F5} ;call up GoTo
    Sleep, 200
    WinWait, Go To, 
    IfWinNotActive, Go To, , WinActivate, Go To, 
    WinWaitActive, Go To, 
    
    ;MsgBox, %tableText%
    Send, %tableText%
    Send, {TAB}
    Send, %entID%{ENTER} ;Call up given window
    Sleep, 250
    ;move control back to advance, jik
    MoveTo("Production Database") 
    
    
    return
}

pullupEntity(entID)
{
    ;pull up entity window
    Send, {ShiftDown}{F4}{ShiftUp}
    Sleep, 25
    Send, %entID%
    Send, {Enter}
}

CopyCell()
{
    clipboard = ;
    Send, {CTRLDOWN}c{CTRLUP}
    ClipWait ;
    ;clean up id from excel
    ;check if you are in excel so the newline can be dropped
    IfWinActive, Excel, , , 
    {
        ;clean up id from excel
        StringTrimRight, clipboard, clipboard, 2
        Send, {Esc}
    }
    cellOfInterest = %clipboard%
    clipboard = ;
    return cellOfInterest
}

Test_for_Nothing(idFound)
{
    ;get window text to check No Entities warning
    Sleep, 1000
    DetectHiddenText, On
    WinGetText, windowCheck 
    ;MsgBox, The text is:`n%windowCheck%
    ;MsgBox 1 idFound = %idFound%

    ;term to check for
    ;windowText = No Entities
    windowText = Looking
    ;Search for windowText in current windows
    IfInString, windowCheck, %windowText%
    {
        idFound := False
        ;Send, {Enter}
        return %idFound%
    }
    return %idFound%
}

Test_for_Multiple(idFound)
{
    ;get window text to check for multiple hits
    WinGetText, windowCheck 
    ;MsgBox, The text is:`n%windowCheck%

    ;term to check for
    windowText = Criteria
    ;Search for windowText in current windows
    IfInString, windowCheck, %windowText%
    {
        ;WinSetTitle, SIS Converter, , NewTitle
        ControlSetText, Static1, User Input Required..., SIS Converter
        MsgBox, 262145,, Please open the correct entity record            
    
        IfMsgBox Cancel 
        {
            Exit
            idFound = False
            return idFound
        }
        ;check if profile was opened or not
        WinGetText, windowCheck
        ;Search for windowText in current windows
        windowText = Custom
        IfNotInString, windowCheck, %windowText%
        {
            Send, {Enter} 
        }
    }
    return idFound
}

;ID capture subroutine
Capture_ID()
{
    
    ;pull up id window
    Send, {F3}{F3}
    Sleep, 250
    businessID := CopyCell()
    ;close the Id window
    Send, {Esc}
    ;close the entity window
    Send, {CTRLDOWN}{F4}{CTRLUP}
    ;close the entity lookup window
    Send, {CTRLDOWN}{F4}{CTRLUP}
    return businessID
}
MoveTo(destination)
{
    ;Move to destination 
    IfWinNotActive, %destination%,
    WinActivate, %destination%, ,
    WinWaitActive, %destination%,
}
+^f::
;MsgBox F started
SetTitleMatchMode, 2

;start in excel
MoveTo("Excel")
businessName := CopyCell()
if not businessName
{
	MsgBox Finished
	exit
}
;prepare adjacent cell to have id pasted
Send, {ShiftDown}{Tab}{ShiftUp}

;move to advance
MoveTo("Production Database")
;open search box, tab to last/org
Send, {SHIFTDOWN}{F4}{SHIFTUP}
Send, {TAB 3}
Send, %businessName%
Send, {ENTER}

MsgBox, 4,, Did you find a matching record?            
IfMsgBox, No 
{
    ;MsgBox No Hit
    Sleep, 50
    Send, {ENTER}
    Sleep, 50
    Send, {ESC}
    ;Send, {CTRLDOWN}{F4}{CTRLUP}
    MoveTo("Excel")
    Send, 0
    Send, {Enter}{Tab}
    Exit
}
;MsgBox Hit
;check if profile was opened or not
WinGetText, windowCheck
;Search for windowText in current windows
windowText = Custom
IfNotInString, windowCheck, %windowText%
{
    Send, {Enter} 
}
businessID := Capture_ID()
;MsgBox, %businessID%

;Move back to Excel
MoveTo("Excel")

;arrow to next column to paste ID
;Send, {LEFT}

;paste ID from clipboard
Send, %businessID%
;MsgBox real thing
Send, {Enter}{Tab}
return
