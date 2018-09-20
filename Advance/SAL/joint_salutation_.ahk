;Context: You need to create joint informal salutations in the SAL windows for two entities
; You are in the MAR window for the entity whose name should come first,
; and the entities should not have any joint informal salutations in the SAL window. You should
; delete them first.
;Instructions: press WindowsKey+j
;Currently the SAL windows stays open after finishing, to have it close itself just delete the first
; semi-colon in the third to last line, before CTRLDOWN

callGoTo(tableText)
{
    SetTitleMatchMode, 2
    Send, {F5} ;call up GoTo
    Sleep, 200
    WinWait, Go To, 
    IfWinNotActive, Go To, , WinActivate, Go To, 
    WinWaitActive, Go To, 
    
    ;MsgBox, %tableText%
    Send, %tableText%{ENTER} ;Call up given window
    Sleep, 250
    ;MsgBox, I've just entered the info
    MoveTo("Production Database") 
    return
}

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

CopyCell()
{
    Sleep, 10
    clipboard = ;
    Send, {CTRLDOWN}c{CTRLUP}
    ClipWait, 0 ;
    cellOfInterest = %clipboard%
    ;clean up id from excel
    ;check if you are in excel so the newline can be dropped
    IfWinActive, Excel, , , 
    {
        ;clean up id from excel
        StringTrimRight, clipboard, clipboard, 2
        Sleep, 10
        cellOfInterest = %clipboard%
        Send, {Esc}
    }
    clipboard = ;
    return cellOfInterest
}

MoveTo(destination)
{
    ;Move to destination 
    ;MsgBox, Moving to %destination%
    IfWinNotActive, %destination%,
    ;MsgBox "2"
    ;#WinActivateForce
    WinActivate, %destination%, ,
    ;MsgBox "4"
    ;WinWaitActive, %destination%
    
}
#j::
SetTitleMatchMode, 1
clipboard = ;Start off empty to allow ClipWait to detect when the text has arrived
MoveTo("Production Database") 
callGoTo("SAL")
Send, {SHIFTDOWN}{F3}{SHIFTUP} ; go to spouse 2 record
Send ^c ; copy spouses informal sal name
ClipWait  ; Wait for the clipboard to contain text.
spouse2 = %clipboard%
clipboard = 
Send, {SHIFTDOWN}{F3}{SHIFTUP} ; go to spouse 1 record
Send ^c ; copy spouses informal sal name
ClipWait  ; Wait for the clipboard to contain text.
spouse1 = %clipboard%
formalSal = %spouse1% and %spouse2%
;MsgBox Control-C copied the following contents to the clipboard:`n`n%formalSal%
;return
clipboard = %formalSal%
Send, {F6}^v
Send, {TAB  2}ji{F8}
Send, {SHIFTDOWN}{F3}{SHIFTUP}
Send, {F6}^v
Send, {TAB  2}ji{F8}
Send, {CTRLDOWN}{F4}{CTRLUP} ;Remove or add the semi-colon before CTRLDOWN to have the SAL window close itself
clipboard =
return
