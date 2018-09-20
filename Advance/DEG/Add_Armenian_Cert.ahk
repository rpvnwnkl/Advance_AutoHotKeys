;
; AutoHotkey Version: 1.x
; Language:       English
; Platform:       Win9x/NT
; Author:         A.N.Other <myemail@nowhere.com>
;
; Script Function:
;	Template script (you can customize this template by editing "ShellNew\Template.ahk" in your Windows folder)
;

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

SetTitleMatchMode, 2
dbName = Production Database
schoolID = 002219
degree = CER
schoolCode = FL
schoolCampus = M
schoolMajor = ARM
classYear = 2017
gradLevel = C
gradType = G

callGoTo(tableText)
{
    global dbName
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
    MoveTo(dbName) 
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
    WinWaitActive, %destination%
    
}

; #IfWinActive, dbName
^+G::
callGoTo("DEG")
;new row
Send, {F6}
Send, %schoolID%

;degree
Send, {TAB}
Send, %degree%

;school
Send, {TAB 3}
Send, %schoolCode%

;campus
Send, {TAB}
Send, %schoolCampus%

;major
Send, {TAB}
Send, %schoolMajor%

;year
Send, {TAB 6}
Send, %classYear%

;level
Send, {TAB 2}
Send, %gradLevel%

;type
Send, {TAB 2}
Send, %gradType%

Send, {F8}
Send, {CtrlDown}{F4}{CtrlUp}

return