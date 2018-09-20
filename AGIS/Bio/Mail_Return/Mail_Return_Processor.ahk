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
;Context: You are inactivating a preferred address
;
;v 1.0
;Pressing windows key + A reloads the script in case it gets wonky

teamID = 246724;
fileName = ;
; Gui, Add, Text,, Please enter the Team ID:
; Gui, Add, Edit, vteamID
; Gui, Add, Button, Default gOK, OK
; Gui, Show
; return

; OK:
; {
;     Gui, Submit
    
; }
fileName = Team%teamID%.csv
;startFile()
#A::
Reload
Return

callGoToBlank(tableText, entID)
{
    SetTitleMatchMode, 2
    Send, {F5} ;call up GoTo
    Sleep, 200
    WinWait, Go To, 
    IfWinNotActive, Go To, , WinActivate, Go To, 
    WinWaitActive, Go To, 
    
    ;MsgBox, %tableText%
    Send, %tableText%
    Send, {TAB}
    Send, %entID%{ENTER} ;Call up given window
    Sleep, 250
    ;MsgBox, I've just entered the info
    WinWait, Production Database, 
    IfWinNotActive, Production Database, , WinActivate, Production Database, 
    WinWaitActive, Production Database, 
    
    
    return
}

pullupEntity(entID)
{
    ;pull up entity window
    Send, {ShiftDown}{F4}{ShiftUp}
    Sleep, 25
    Send, %entID%
    Send, {Enter}
}

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

inactivateAddress(comment)
{
    global ;
    ;go to type field
    Send, {AltDown}Y{AltUp}
    addrType := CopyCell()
    If addrType in A,H
    {
        Send, P
    } Else If addrType in B,D
    {
        Send, Q
        addrLinked = False
        MsgBox, 4, EMPT Info, Is the Business Address Linked to an employment record?
        IfMsgBox, Yes
        {
            addrLinked = True
        }

    } Else If addrType in S
    {
        Send, T
    } Else If addrType in C,L
    {
        Send, U
    } Else
    {
        MsgBox, Unknown Address Type
        return
    }

    ;go to status field
    Send, {TAB}
    addrStatus := CopyCell()
    If addrStatus in I,N,K
    {
        ;changes have been made and no more are necessary
        MsgBox, This address has already been updated
        return
    } Else
    {
        ;mark inactive
        Send, I
    }

    /*
    ;go to preferred box
    Send, {TAB 7}
    Send, {Space}
    */
    ;capture start month
    Send, {TAB}
    Send, {ShiftDown}{Right 2}{ShiftUp}
    modM := CopyCell()

    ;capture start day
    Send, {TAB}
    Send, {ShiftDown}{Right 2}{ShiftUp}
    modD := CopyCell()

    ;capture start month
    Send, {TAB}
    Send, {ShiftDown}{Right 4}{ShiftUp}
    modY := CopyCell()


    If addrLinked
    {
        ;skip linked box
        Send, {Tab 5}
    } Else
    {
        ;go to preferred box
        Send, {TAB 4}
    }
    ;uncheck preferred
    Send, {Space}
    
    ;capture source code
    Send, {TAB}
    modCode := CopyCell()

    ;go to comment field
    Send, {TAB 3}
    ;select comment field and save to variable
    modComment := CopyCell()
    IfEqual, modComment, 
    {
        Send, %comment%
    } Else
    {
        Send, {Right}; %comment%
    }

    ;now save, return
    Send, {F8}

    ;for telephone record
    Send, {ENTER}
    ;for business info
    if addrLinked
    {
        Send, {ENTER}
    }
    ;close two addr windows
    Sleep, 100
    Send, {CtrlDown}{F4}{CtrlUp}
    Sleep, 100
    Send, {CtrlDown}{F4}{CtrlUp}

}

startFile()
{
    global ;
    Loop, Read, %fileName%, %fileName%
    {
        MsgBox, starting file loop
        IfNotInString, A_LoopReadLine, `n
        {
            ;an empty file
            MsgBox, Empty File
            textToLog = EntityId, Record Status, Addr. Start Date, Source Code, Comment`n
            FileAppend, %textToLog%, %fileName%
        }
    }
}

writeToFile()
{
    global ;
    ;fileName = Team%teamID%.csv
    textToLog = %entityID%, %recordStatus%, %modM%/%modD%/%modY%, %modCode%, %modComment%`n
    FileAppend, %textToLog%, %fileName%
}
;The middle mouse button launches the script
;^MButton::
+^G::
DetectHiddenWindows, On 
DetectHiddenText, On 
SetTitleMatchMode 2 
SetTitleMatchMode Slow
Sleep, 100

comment = Team %teamID%
;LogFileName = Team241508.txt
modM = ;
modD = ;
modY = ;
modCode = ;
modComment = ;

;Copy highlighted ID
entityID := CopyCell()

;move to Advance
WinWait, Production Database, 
IfWinNotActive, Production Database, , WinActivate, Production Database, 
WinWaitActive, Production Database, 
Sleep, 100



/*
Sleep, 500
;get modified date
WinGetText, windowText, A, , 
MsgBox, %windowText%
searchText = Modified:
StringGetPos, ModLocation, windowText, %searchText%
MsgBox, %ModLocation%
*/
;pull up address
callVar = ADDR
callGoToBlank(callVar, entityID)

MsgBox, 4, Confirm Update, Update Address?
IfMsgBox, Yes
{
    ;mark incative
    ;pull up preferred/first addr line
    Send, {ENTER}
    inactivateAddress(comment)
    
} Else
{
    exit
}

;refresh to view new bio screen
Sleep, 200
pullupEntity(entityID)
Sleep, 200
; Send, {CtrlDown}U{CtrlUp}
;Send, {F7}
MsgBox, 4, Entity Status, Is the status still Active?
IfMsgBox, Yes
{
    recordStatus = Active
    ;pull up tel window
    ; callVar = TEL
    ; callGoToBlank(callVar, entityID)

    ;return to excel and mark yellow
    WinWait, Excel, 
    IfWinNotActive, Excel, , WinActivate, Excel, 
    WinWaitActive, Excel,

    ;color the cell orange
    Send, {AltDown}HH{AltUp}
    Send, {Right 7}
    Send, {Enter}
    ;next cell
    Send, {Enter}

} Else
{
    MsgBox, Mark the record Lost or find a new address
    recordStatus = Lost
    ;pull up tel window
    ; callVar = TEL
    ; callGoToBlank(callVar, entityID)

    ;return to excel and mark yellow
    WinWait, Excel, 
    IfWinNotActive, Excel, , WinActivate, Excel, 
    WinWaitActive, Excel,

    ;color the cell red
    Send, {AltDown}HH{AltUp}
    Send, {Down 6}{Right}
    Send, {Enter}
    ;next cell
    Send, {Enter}
}
;MsgBox, About to write file
writeToFile()
return
/*
;close previous windows
Send, {CtrlDown}{F4 3}{CtrlUp}
;open Entity search window
;Send, {SHIFTDOWN}{F4}{SHIFTUP}
;Sleep, 100
pullupEntity(entityID)

callVar = DEG
callGoToBlank(callVar, entityID)

callVar = REC
callGoToBlank(callVar, entityID)

;Paste id into first text box
;Send, %entityID%

;Sleep, 100
;Send, {ENTER}
;Sleep, 100
*/
return