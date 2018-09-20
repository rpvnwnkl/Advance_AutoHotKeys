;Context: You have an ID highlighted and would like to search it in Advance
;	
;	the script copies ID and moves over to advance 
;	an entity search window is opened and the ID pasted in
;	

;Instructions: have advance open, select the ID you'd like to search 
;	execute this script by pressing the middle mouse button
;	
;v 1.0
;Pressing windows key + A reloads the script in case it gets wonky
ogClipContents = %clipboard%

SetTitleMatchMode, 2
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

MoveTo(destination)
{
    ;Move to destination 
    ;MsgBox, Moving to %destination%
    IfWinNotActive, %destination%,
    ;MsgBox "2"
    ;#WinActivateForce
    WinActivate, %destination%, ,
    ;MsgBox "4"
    WinWaitActive, %destination%
    
}

callProsA(prosID)
{
    SetTitleMatchMode, 2
    Send, {F5} ;call up GoTo
    Sleep, 200
    WinWait, Go To, 
    IfWinNotActive, Go To, , WinActivate, Go To, 
    WinWaitActive, Go To, 
    
    ;MsgBox, %tableText%
    Send, PROSA
    Send, {TAB 2}
    Send, %prosID%{ENTER} ;Call up given window
    Sleep, 250
    ;move control back to advance, jik
    MoveTo("Production Database") 
    
    
    return
}
;copy entity ID
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
} else 
{
    ;move to Advance
    WinWait, Production Database,
    IfWinNotActive, Production Database, , WinActivate, Production Database,
    WinWaitActive, Production Database,
    Sleep, 100

    ;open PROSA window
    callProsA(entityID)
    Sleep, 100

}

clipboard = %ogClipContents%
return