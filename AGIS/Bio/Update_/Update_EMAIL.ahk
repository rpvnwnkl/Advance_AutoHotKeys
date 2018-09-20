
;v 1.0
SetTitleMatchMode, 2

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

CopyAndCheck()
{
    ;Copy highlighted ID
    ;global
    entityID := CopyCell()

    IfEqual, entityID, 
    {
        ;MsgBox, Done
        return
    }

    IfWinNotExist, Production Database
    {
        MsgBox, Advance isn't open
        return
    }

    ;move to Advance
    MoveTo("Production Database")
    Sleep, 100 
    return entityID   
}

; #IfWinActive, Production Database
; ^+U::
;open windows

; callGoTo("TEL")
; callGoTo("ADDR")
; callGoTo("EMPT")
; callGoTo("EMAIL")

; return

; #IfWinActive, Production Database
; Esc::
; Send, {CtrlDown}{F4}{CtrlUp}
; return

; #IfWinActive, Production Database
; ^N::
; Send, {ShiftDown}{F4}{ShiftUp}
; Send, {TAB 3}
; return

#E::
;copy id
entID := CopyCell()
;copy email
Send, {RIGHT}
emailAddr := CopyCell()
;move to advance
MoveTo("Production Database")
;open email dialog
callGoToBlank("EMAIL", entID)
;enter email
Send, {F6}
Send, %emailAddr%
;preferred box
Send, {TAB} ;move
Send, {Space}
;type B for business
Send, {TAB 3} ;move
Send, B
;source code
Send, {TAB 3} ;move
Send, BIO
;comment field
Send, {TAB 6} ;move
Send, Team 248541
;save
Send, {F8}
;close
Send, {ENTER} ;for rogue dialogues
Send, {CtrlDown}{F4}{CtrlUp}
;Back to Xcel
MoveTo("Excel")
;color box
Send, {Left}{Alt}HH{Right 5}{Enter}
Send, {Down}

return

