SendMode, Input
; dbName = ADVTRN
createGUI()
{
    global dbName
    global dbVal
    Gui, Font, S16 CDefault Bold Italic, Magneto
    Gui, Add, Text, x12 y19 w320 h40 , Telefund Processor v0.1

    Gui, Font, S12 Cblack Italic, Times
    Gui, Add, Picture, x52 y59 w200 h170 , C:\Users\mmills05\OneDrive - Tufts\AGIS\Telefund\Cookie mouth open.png

    Gui, Font, S14 Cblack norm, consolas
    Gui, Add, Text, x22 y232 w130 h30 , Database:
    Gui, Font, S12 Cblack norm, courier new
    Gui, Add, DropDownList, x122 y232 w150 h30 R2 vdbVal, Production||Training 
    ; Generated using SmartGUI Creator 4.0
    Gui, Show, x956 y107 h278 w306, Telefund Processor
    Gui, +AlwaysOnTop
    Return

    ; Generated using SmartGUI Creator 4.0
    ; Gui, Show, x961 y119 h228 w307, Telefund Processor
    ; Return

    GuiDropFiles:
        gui, submit, nohide
        ; msgBox, %dbVal%
        IfEqual, dbVal, Production
        {
            dbName = Production Database
        }else
        {
            dbName = ADVTRN
        }
        IfWinNotExist, %dbName%
        {
            MsgBox, %dbName% isn't open.

        } else 
        {
            StringSplit, F, A_GuiEvent, `n
            Loop, %F0%
            {
                currVal := F%A_Index%
                ; currVal = %currVal%
                ; MsgBox, %currVal%
                SplitPath, currVal,,fileDir,,fileDate
                SetWorkingDir, %fileDir%

                ; msgbox, %A_WorkingDir%
                processCSV(currVal)
            }
        }
    ;  FileRead, Text, %F1%
    ;  GuiControl,, MyEdit, %F1%
        Return

    GuiClose:
    GuiEscape:
     ExitApp
}
createGUI()
moveTo(destination)
{
    ;Move to destination 
    ;MsgBox, Moving to %destination%
    IfWinNotActive, %destination%,
    ;MsgBox "2"
    ;#WinActivateForce
    WinActivate, %destination%, ,
    ;MsgBox "4"
    WinWaitActive, %destination%
    return
    
}

CopyCell()
{
    Sleep, 10
    clipboard = ;
    Send, {CTRLDOWN}c{CTRLUP}
    ClipWait, 0 ;
    cellOfInterest = %clipboard%
    ;clean up id from excel
    ;check if you are in excel so the newline can be dropped
    IfWinActive, Excel, , , 
    {
        ;clean up id from excel
        StringTrimRight, clipboard, clipboard, 2
        Sleep, 10
        cellOfInterest = %clipboard%
        Send, {Esc}
    }
    clipboard = ;
    return cellOfInterest
}

pullupEntity(entID)
{
    ;pull up entity window
    Send, {ShiftDown}{F4}{ShiftUp}
    Sleep, 25
    Send, %entID%
    Send, {Enter}
}

callGoToBlank(tableText, entID)
{
    global dbName
    SetTitleMatchMode, 2
    Send, {F5} ;call up GoTo
    Sleep, 200
    WinWait, Go To, 
    IfWinNotActive, Go To, , WinActivate, Go To, 
    WinWaitActive, Go To, 
    
    ;MsgBox, %tableText%
    Send, %tableText%
    Send, {TAB}
    ; MsgBox, %entID%
    Send, %entID%{ENTER} ;Call up given window
    Sleep, 250
    ;move control back to advance, jik
    MoveTo(dbName) 
    
    
    return
}

MarkNoTelefund(callID)
{
    ; pullupEntity(callID)
    callGoToBlank("MAILC", callID)
    MsgBox, No Telefund
    Send, {TAB}
    currVal := CopyCell()
    if not currVal
    {
        ;it's empty
        ;all other values are more exclusive
        Send, 3   
    } 
    ;Tab to No Call field
    Send, {TAB 11}
    Send, Y
    ;Tab to comment
    Send, {TAB 3}
    CommentCRL()
}

MarkNoSolicit(callID)
{
    ; pullupEntity(callID)
    callGoToBlank("MAILC", callID)
    MsgBox, No Solicitation
    Send, {TAB}
    currVal := CopyCell()
    IfNotEqual, currVal, 1
    {
        ;it's not 1
        ;change to 1 as it's most selective
        Send, 1
    }
    ;ensure proper controls are set
    ;move to S Mail field
    Send, {TAB 6}
    Send, Y
    ;move to S email
    Send, {TAB 3}
    Send, Y
    ;move to S Calls
    Send, {TAB 2}
    Send, Y

    ;check for comment
    Send, {TAB 3}
    CommentCRL()
    return
    
}

MarkFullWithhold(callID)
{
    ; pullupEntity(callID)
    callGoToBlank("MAILC", callID)
    MsgBox, Full Withhold 
    Send, {TAB}
    currVal := CopyCell()
    IfNotEqual, currVal, 1
    {
        ;it's not 1
        ;change to 1 as it's most selective
        Send, 1
    }
    ;ensure proper controls are set
    ;move to S Mail field
    Send, {TAB 6}
    Send, Y
    ;move to S email
    Send, {TAB 3}
    Send, Y
    ;move to S Calls
    Send, {TAB 2}
    Send, Y

    ;check for comment
    Send, {TAB 3}
    CommentCRL()
    return
    
}
CommentCRL()
{
    global callDate
    ;select All
    Send, {Alt}
    Send, el
    currVal := CopyCell()
    IfEqual, currVal,
    {
        ;it's empty
        Send, Per CRL %callDate%
    }Else
    {
        ;it's not empty
        ;does it contain CRL already?
        IfNotInString, currVal, CRL
        {
            Send, {RIGHT}
            Send, `; Per CRL %callDate%
        }
        
    }
}

JustSave()
{
    Send, {F8}
    ;close mailc
    Send, {CtrlDown}{F4}{CtrlUp}
    Sleep, 50
}

SavePrompt()
{
    MsgBox, 4, Save Changes, Save Modifications?
    IfMsgBox, Yes
    {
        ;save mailc
        Send, {F8}
        ;close mailc
        Send, {CtrlDown}{F4}{CtrlUp}
        ;close entity window
        Send, {CtrlDown}{F4}{CtrlUp}
    } Else
    {
        Send, {F7}
        ;close mailc
        Send, {CtrlDown}{F4}{CtrlUp}
        ;close entity window
        Send, {CtrlDown}{F4}{CtrlUp}
    }
    return
}

ensureAdvance()
{
    global dbName
    ;move to advance
    IfWinNotExist, %dbName%
    {
        MsgBox, Advance isn't open.
        exit 
    }
    WinWait, %dbName%, 
    IfWinNotActive, %dbName%, , WinActivate, %dbName%, 
    WinWaitActive, %dbName%, 
    Sleep, 50
}

updateMailControls(idNum, callResult, callComment)
{
    global dbVal
    ensureAdvance()
    ;Call result codes
    DoNotCall = Do Not Call
    RemoveList = Remove From List
    changeMade = False
    ; msgBox, %idNum
    ;determine call result
    IfEqual, callResult, %DoNotCall%
    {
        MarkNoTelefund(idNum)
        FileAppend, 
        (
        %A_TAB%Marking ID# %idNum% for No Telefund`n
        ), %dbVal%_processorLog.txt
        changeMade = True
    } Else IfEqual, callResult, %RemoveList%
    {
        MarkNoSolicit(idNum)
        FileAppend, 
        (
        %A_TAB%Marking ID# %idNum% for No Solicitation`n
        ), %dbVal%_processorLog.txt
        changeMade = True
    }

    ;check for comment
    IfNotEqual, callComment,
    {
        FileAppend, 
        (
        %A_TAB%Comment text: %callComment%`n
        ), %dbVal%_processorLog.txt
        ;a comment exists
        MsgBox, 1, Comment in CRL, `nComment: `n`n%callComment%
        IfMsgBox, Cancel
        {
            FileAppend, 
            (
            %A_TAB%Previous update may not have been saved.`n
            ), %dbVal%_processorLog.txt
            Exit
        }
    }
    IfEqual, changeMade, True
    {
        JustSave()
    }
}

processCSV(fName)
{
    updateMade = False
    global dbName
    global dbVal
    global notADVid = 0
    global realADVid = 0
    global callDate
    ;extracting calldate from filename
    SplitPath, fName, , , , callDate,

    Loop, read, %dbVal%_completedCRLs.txt
    {
        IfInString, A_LoopReadLine, %callDate%
        MsgBox, The results from %callDate% have already been processed.
        return
    }
    Loop, read, %fName%
    {
        
        LineNumber = %A_Index%
        if (LineNumber = 1)
        {
            
            FileAppend,
            (
            `nStarting call results from %callDate%...`n
            ), %dbVal%_processorLog.txt
        }
        ; msgBox, %A_LoopReadLine%
        StringSplit, thisLine, A_LoopReadLine, `,, %A_Space%"
        if (LineNumber > 1) 
        {
            personNum := thisLine1
            idNum := thisLine2
            PageNum := thisLine3
            CallResult := thisLine4
            comments := thisLine5
            dateCalled := thisLine6
            ; msgBox, %idNum%
            ; StringReplace, idNum, idNum, `",
            ; msgBox, %idNum%
            StringLen, idLength, idNum
            if (idLength > 10)
            {
                notADVid += 1
                continue
            } else
            {
                realADVid += 1
                IfEqual, CallResult, Deceased
                {
                    FileAppend,
                    (
                    Deceased, %dateCalled%, %idNum%, %comments%`n
                    ), %dbVal%_deceasedLog.csv
                } else
                {
                    updateMailControls(idNum, CallResult, comments)
                    updateMade = True
                }
            }
        }
        ; MsgBox, personNum = %personNum%`nidNum = %idNum%`nPageNum = %PageNum%`nCallResult = %CallResult%`ncomment = %comments%
        ; Loop, parse, A_LoopReadLine, `, 
        ; {
        ;     MsgBox, 4,, Field %LineNumber%-%A_Index% is:`n%A_LoopField%`n`nContinue?
        ;     IfMsgBox, No
        ;         return
        ; }
    }
    FileAppend,
    (
    
    Advance IDs processed: %realADVid%`n
    Other IDs skipped during processing: %notADVid%`n
    ), %dbVal%_processorLog.txt
    FileAppend,
    (
        %callDate%`n
    ), %dbVal%_completedCRLs.txt

    msgBox, Log updated with call results from %callDate%.
}