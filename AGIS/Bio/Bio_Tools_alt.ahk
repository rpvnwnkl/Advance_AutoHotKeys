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

;look up phone number with Twilio
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

;features for updating CTS
#Include, windowPresent.ahk
#Include, Bio\updateCTSInfo.ahk
#IfWinActive, Production Database
^1::
{
    If windowPresent("Affiliation")
    {
        updateCTSInfo("1701")
    }
    return
}
#IfWinActive, Production Database
^2::
{
    If windowPresent("Affiliation")
    {
        updateCTSInfo("1702")
    }
    return
}
#IfWinActive, Production Database
^3::
{
    If windowPresent("Affiliation")
    {
        updateCTSInfo("1703")
    }
    return
}
#IfWinActive, Production Database
^4::
{
    If windowPresent("Affiliation")
    {
        updateCTSInfo("1704")
    }
    return
}
#IfWinActive, Production Database
^5::
{
    If windowPresent("Affiliation")
    {
        updateCTSInfo("1705")
    }
    return
}
#IfWinActive, Production Database
^6::
{
    If windowPresent("Affiliation")
    {
        updateCTSInfo("1706")
    }
    return
}
;##############################################
;#### Instant Address
;##############################################

;move Instant Address
<!3::
{
    ; msgbox, hi
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
        Send, {AltDown}{3}
        Send, {AltUp}
    }
    return
}

;##############################################
;#### TEAM 
;##############################################

;Save TEAM Request as WIP
#Include Team\saveTeamRequest.ahk
#IfWinActive, Team
#T::
saveTeamRequest()
return

;close active Team window
#IfWinActive, Team Work Request
Esc::
{
    Send, {AltDown}{F4}
    Send, {AltUp}
    return
}

;Open various windows from TEAM
#Include TEAM\Open.ahk
;Telephone
#IfWinActive, Team
^+T::
{
    Open("TEL")
    return
}
;Address
#IfWinActive, Team
^+A::
{
    Open("ADDR")
    return
}
;Email
#IfWinActive, Team
^+M::
{
    Open("EMAIL")
    return
}
;Employment
#IfWinActive, Team
^+E::
{
    Open("EMPT")
    return
}

