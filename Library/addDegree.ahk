;
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetTitleMatchMode, 2

#Include callGoToBlank.ahk
#Include moveTo.ahk

addDegree(entID, instID, degCode, schoolCode, campusCode, majorCode, classYear, gradLevel, gradType, honAlum, localAlum)
{
    ; instID
    ; degCode
    ; schoolCode
    ; campusCode
    ; majorCode
    ; classYear
    ; gradLevel
    ; gradType
    ; honAlum
    ; localAlum

    global dbName
    callGoToBlank("DEG", entID)
    ;new row
    Send, {F6}
    Send, %instID%
    Sleep, 15
    ;degree
    Send, {TAB}
    Send, %degCode%

    ;school
    Send, {TAB 3}
    Send, %schoolCode%

    ;campus
    Send, {TAB}
    Send, %campusCode%

    ;major
    Send, {TAB}
    Send, %majorCode%

    ;comment
    Send, {TAB 5}
    Send, BIO 

    ;year
    Send, {TAB 1}
    Send, %classYear%

    ;level
    Send, {TAB 2}
    Send, %gradLevel%

    ;type
    Send, {TAB 2}
    Send, %gradType%

    ; localAlum
    Send, {TAB}
    IfEqual, localAlum, Y
    {
        Send, {Space}
    }

    ; honAlum
    Send, {TAB}
    IfEqual, honAlum, Y
    {
        Send, {Space}
    }

    ; save and close
    Send, {F8}
    ; Send, {CtrlDown}{F4}{CtrlUp}

    return
}