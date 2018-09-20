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

SetTitleMatchMode, 2
MoveTo("Production Database") 
callGoTo("SAL")
Send, {SHIFTDOWN}{F3}{SHIFTUP} ; go to spouse 2 record
spouse2 := copyCell()
Send, {SHIFTDOWN}{F3}{SHIFTUP} ; go to spouse 1 record
spouse1 := copyCell()
formalSal = %spouse1% and %spouse2%
Send, {F6}
Send, %formalSal%
Send, {TAB  2}
Send, ji
Send, {F8}
Sleep, 100
Send, {SHIFTDOWN}{F3}{SHIFTUP}
Send, {F6}
Send, %formalSal%
Send, {TAB  2}
Send, ji
Send, {F8}
Sleep, 100
Send, {CTRLDOWN}{F4}{CTRLUP} ;Remove or add the semi-colon before CTRLDOWN to have the SAL window close itself
return
