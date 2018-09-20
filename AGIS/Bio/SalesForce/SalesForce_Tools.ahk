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

ensureAdvance()
{
    ;move to advance
    WinWait, Production Database, 
    IfWinNotActive, Production Database, , WinActivate, Production Database, 
    WinWaitActive, Production Database, 
    Sleep, 50
}

#A::
Reload
Return

#P::
SetTitleMatchMode, 2
ensureAdvance()
;pull up menu, key to Edit and then Post
Send, {AltDown}{AltUp}et
return 

;Discard
#D::
SetTitleMatchMode, 2
ensureAdvance()
;pull up menu, key to Edit and then Post
Send, {AltDown}er{AltUp}
;choose duplicate and exit
Send, {Down}
Send, {Tab 2}
Send, O
return

;move apt up
#M::
Run, Move_Apt_SF_live.ahk,,,
return

;switch line one two
#N::
Run, Swap_1_2_SF_live.ahk,,,
return

;change to modify, validate, and post
#I::
Run, Mod_Validate_live.ahk,,,
return