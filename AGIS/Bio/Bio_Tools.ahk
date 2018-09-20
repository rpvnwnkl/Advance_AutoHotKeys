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
openCount = 0
#Include %A_ScriptDir%\..\..\Library
#Include callGoTo.ahk
#Include moveTo.ahk

;Add informal Salutation
; +^i::
; Run, informal_live.ahk, ADV\SAL
; Return
#Include newEntitySalutation.ahk
+^i::
newEntitySalutation()
return

;Add joint Salutation
; +^j::
; Run, joint_salutation_live.ahk, ADV\SAL
; Return
#Include createMarriedSals.ahk
+^j::
createMarriedSals()
return

;Open ID in Advance
; MButton::
; Run, Open_ADV_ID_live.ahk, Open_\
; return
#Include Open_Adv_ID.ahk
MButton::
{
    openCount += 1
    Open_ADV_ID()
    return
}

; #IfWinActive Bio_Tools
!^i::MsgBox, %openCount%

;Save TEAM Request as WIP
#T::
Run, Save_Request_live.ahk, TEAM\
return

;look up phone number
#C::
{
    Run, C_Check_Number_Live.ahk, ..\..\Advance\TEL
    return
}

;macros specific to advance

;open typical bio windows
#IfWinActive, Production Database
^+U::
callGoTo("TEL")
callGoTo("ADDR")
callGoTo("EMPT")
callGoTo("EMAIL")
return

;open deceased windows
#IfWinActive, Production Database
^+D::
callGoTo("ADDR")
callGoTo("TEL")
callGoTo("EMAIL")
return

;close current window with ESC key
#IfWinActive, Production Database
Esc::
Send, {CtrlDown}{F4}{CtrlUp}
return

;open new entity search box and tab to last name
#IfWinActive, Production Database
^N::
Send, {ShiftDown}{F4}{ShiftUp}
Send, {TAB 3}
return