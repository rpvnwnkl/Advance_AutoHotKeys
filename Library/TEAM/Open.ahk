#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
; SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetTitleMatchMode, 2
#Include ..\Library
#INclude CallGoToBlank.ahk
#Include CopyAndCheck.ahk

Open(windowCode)
{
    entID := CopyAndCheck()
    callGoToBlank(windowCode, entID)
    return
}
;     IfEqual, windowLetter, T
;     {
;         ;open TEL window
;         callGoToBlank("TEL", entID)
;         return
;     }
;     IfEqual, windowLetter, M
;     {
;         ;open EMAIL window
;         callGoToBlank("EMAIL", entID)
;         return
;     }
;     IfEqual, windowLetter, E
;     {
;         ;open EMPT window
;         callGoToBlank("EMPT", entID)
;         return
;     }
;     IfEqual, windowLetter, A
;     {
;         ;open ADDR window
;         callGoToBlank("ADDR", entID)
;         return
;     }
; }