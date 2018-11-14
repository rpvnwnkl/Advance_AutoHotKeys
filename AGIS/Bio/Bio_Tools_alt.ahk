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
#Include newEntitySalutation.ahk
+^i::
newEntitySalutation()
return

;Add joint Salutation
#Include createMarriedSals.ahk
+^j::
createMarriedSals()
return

;Open ID in Advance
#Include Open_Adv_ID.ahk
^LButton::
MButton::
{
    openCount += 1
    Open_ADV_ID()
    return
}

; #IfWinActive Bio_Tools
!^i::MsgBox, %openCount%

;Save TEAM Request as WIP
#Include Team\saveTeamRequest.ahk
#T::
saveTeamRequest()
return

;look up phone number
#Include TEL\checkTelephone.ahk
#IfWinActive, Production Database
#C::
{
    checkTelephone()
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
callGoTo("ENT")
callGoTo("ADDR")
callGoTo("TEL")
callGoTo("EMAIL")
callGoTo("ECON")
callGoTo("AFFIL")
callGoTo("BDET")
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

today()
{
    FormatTime, month,,MM
    FormatTime, day,,dd
    FormatTime, year,,yyyy
    today = %month%/%day%/%year%
    return today
}

#IfWinActive, Production Database
^LButton::
+^T::
today := today()
Send, %today%
return

; #Include, windowPresent.ahk
#Include, Bio\updateBioInfo.ahk
#IfWinActive, Production Database
^K::
{
    updateBioInfo("K")
    return
}

#IfWinActive, Production Database
^P::
{
    updateBioInfo("P")
    return
}

;move Instant Address
<!3::
{
    MoveTo(dbName)
    Send, {AltDown}{3}
    Sleep, 10
    Send, {AltUp}
    Loop, 100
    {
        IfWinExist, Instant Address Validation
        {
            Sleep, 50
            ; msgbox, hello
            WinActivate, 
            WinMove, Instant Address Validation, , 350, 200, 
            return
        }
        Sleep, 10
    }
    return
}