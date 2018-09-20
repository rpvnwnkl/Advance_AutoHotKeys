#Include waitForWindow.ahk
#Include closeWindows.ahk

createEntity(prefix, firstName, middleName, lastName, suffix, proSuffix, recYear, recSchool, recType, recStatus, comment)
{
    global dbName
    ;pull up new entity window
    Send, {ShiftDown}{F4}{ShiftUp}
    Sleep, 25
    Send, {F6}

    ;###########################
    ;start with name
    Send, {AltDown}L{AltUp}
    Send, %lastName%
    Send, {Tab}
    ;####################
    ;DOES NOT ACCOUNT FOR INITIAL FIRST NAME
    Send, %firstName%
    Send, {Tab}
    Send, %middleName%
    Send, {Tab}
    Send, %prefix%
    Send, {Tab}
    Send, %suffix%
    Send, {Tab}
    Send, %proSuffix%

    ;############################
    ;Now record type info
    ;year, school + type
    Send, {Tab 2}
    Send, %recYear%
    Send, {Tab}
    Send, %recSchool%
    Send, {Tab}
    Send, %recType%
    Send, {Tab}
    Send, %recStatus%

    ;###############################
    ;add comment
    Send, {AltDown}m{AltUp}
    Send, %comment%
    
    ;########################
    ;Now, save and review
    Send, {F8}

    IfNotEqual, proSuffix,
    {
        ;need to ok on validate entity popup
        WinWaitActive, Validate Entity
        Send, {Enter}
        ; Msgbox, just closed validate box?
    }

    ;check for duplicate entities window
    moveTo(dbName)
    dupeExists := waitForWindow("Duplicate Entities", 25)
    ;prompt to save anyway/cancel
    IfEqual, dupeExists, 1
    {
        MsgBox, 4097, ,Possible Dupe found, Save Anyway?
            IfMsgBox, Ok
            {
                ;save
                moveTo(dbName)
                ;close dupe box
                Send, {Enter}
                ;submit on save anyway box
                Send, {Enter}

                ;check for validate window again
                IfNotEqual, proSuffix,
                {
                    ;need to ok on validate entity popup
                    WinWaitActive, Validate Entity
                    Send, {Enter}
                    ; Msgbox, just closed validate box?
                }
                ;now continue to save regular stuff
            } else
            {
                ;don't save
                moveTo(dbName)
                Send, {Enter}
                Send, N
                closeWindows()
                return 0
            }
    } Else
    {
        ; Msgbox, No dupe found
    }
    
    WinWaitActive, Primary Entity
    IfWinActive, Primary Entity
    {
        Send, {Enter}
    }
    
    ;now no address box
    IfWinActive, Primary Entity
    {
        Send, {Enter}
    }

    ;retrieve ID
    moveTo(dbName)
    entityReady := waitForWindow("Entity - (")
    ; msgBox, %entityReady%
    IfEqual, entityReady, 1
    {
        ;gathering new ID number
        ;get window text to get ID
        WinGetText, entityWindow 
        Sleep, 200
        ;while we could find the position of the id num
        ;we know where it is for now, so we will hard code it
        StringMid, idNum, entityWindow, 15, 6
        return %idNum%
    } Else
    {
        MsgBox, check for errors, Id not captured
        return 0
    }
}