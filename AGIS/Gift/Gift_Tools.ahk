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

#A::
Reload
Return

;Add informal Salutation
^MButton::
Run, Open_Gift_Live.ahk, ADV\Gift
Return

;Open ID in Advance
MButton::
Run, Open_ADV_ID_live.ahk, TEAM\
return

;Save TEAM Request as WIP
#T::
Run, Save_Request_live.ahk, TEAM\
return