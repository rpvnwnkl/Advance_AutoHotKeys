#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetTitleMatchMode, 2
#Include %A_ScriptDir%\..\..\..\Library
#Include Copycell.ahk
#INclude CallGoToBlank.ahk
#INclude MoveTo.ahk
#Include CopyAndCheck.ahk
dbName = Production

; CopyCell()
; {
;     Sleep, 10
;     clipboard = ;
;     Send, {CTRLDOWN}c{CTRLUP}
;     ClipWait, 0 ;
;     cellOfInterest = %clipboard%
;     ;clean up id from excel
;     ;check if you are in excel so the newline can be dropped
;     IfWinActive, Excel, , , 
;     {
;         ;clean up id from excel
;         StringTrimRight, clipboard, clipboard, 2
;         Sleep, 10
;         cellOfInterest = %clipboard%
;         Send, {Esc}
;     }
;     clipboard = ;
;     return cellOfInterest
; }
; callGoToBlank(tableText, entID)
; {
;     SetTitleMatchMode, 2
;     Send, {F5} ;call up GoTo
;     Sleep, 200
;     WinWait, Go To, 
;     IfWinNotActive, Go To, , WinActivate, Go To, 
;     WinWaitActive, Go To, 
    
;     ;MsgBox, %tableText%
;     Send, %tableText%
;     Send, {TAB}
;     Send, %entID%{ENTER} ;Call up given window
;     Sleep, 250
;     ;move control back to advance, jik
;     MoveTo("Production Database") 
    
    
;     return
; }
; MoveTo(destination)
; {
;     ;Move to destination 
;     ;MsgBox, Moving to %destination%
;     IfWinNotActive, %destination%,
;     ;MsgBox "2"
;     ;#WinActivateForce
;     WinActivate, %destination%, ,
;     ;MsgBox "4"
;     ;WinWaitActive, %destination%
    
; }

; CopyAndCheck()
; {
;     ;Copy highlighted ID
;     global
;     entityID := CopyCell()

;     IfEqual, entityID, 
;     {
;         ;MsgBox, Done
;         return
;     }

;     IfWinNotExist, Production Database
;     {
;         MsgBox, Advance isn't open
;         return
;     }

;     ;move to Advance
;     MoveTo("Production Database")
;     Sleep, 100 
;     return entityID   
; }

^+T::
entID := CopyAndCheck()
;open TEL window
callGoToBlank("TEL", entID)
return

^+M::
entID := CopyAndCheck()
;open EMAIL window
callGoToBlank("EMAIL", entID)
return

^+E::
entID := CopyAndCheck()
;open EMPT window
callGoToBlank("EMPT", entID)
return

^+A::
entID := CopyAndCheck()
;open ADDR window
callGoToBlank("ADDR", entID)
return
