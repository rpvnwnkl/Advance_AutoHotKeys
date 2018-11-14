#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetTitleMatchMode, 2

dbName = Production Database
#Include %A_ScriptDir%\..\..\..\Library

#Include MoveTo.ahk
#Include pullUpEntity.ahk
#Include CallGoToBlank.ahk
#Include CallGoTo.ahk
#Include quietClick.ahk
#Include waitForWindow.ahk
#Include closeBackgroundWindow.ahk

moveAdvWindow(x,y)
{
    MoveTo(dbName)
    ControlGet, childWin, Hwnd,, FNWND31201, Production
    ; Msgbox, %childWin%
    ControlMove, , %x%, %y%, , , ahk_id %childWin% 
    return
}

;Entity Record
; callGoToBlank("PROF", "681810")

!p::
{
    msgbox, hi
    
    ;prospect record
    ;first summary view
    Send, {F7}
    quietClick("Button1", dbName, "Custom Tab Profile", 0)
    ;now info button
    Send, {F7}
    quietClick("Button1", dbName, "Prospect Tracking Summary", 0)
    ;fix delay here
    waitForWindow("Prospect -(", 5)
    moveAdvWindow(815, 615)

    ;close prospect tracking window
    closeBackgroundWindow(1)
    
    ;giving summary
    callGoTo("GSUM")
    moveAdvWindow(815, 95)
    
    ;contact reports
    callGoTo("CRPTS")
    moveAdvWindow(55, 585)
    
    Return
}



; Sleep, 50
; ControlGet, childWin, Hwnd,, FNWND31201, Production
; Msgbox, %childWin%
; ControlMove, , 800, 75, , , ahk_id %childWin% 