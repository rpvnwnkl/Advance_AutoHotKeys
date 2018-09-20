#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetTitleMatchMode, 2
#Include ..\..\..\Library
#Include CopyCell.ahk
#Include callGoToBlank.ahk
#Include pullUpEntity.ahk
#Include moveTo.ahk
#Include ensureAdvance.ahk

createGUI()
{
    global dbName
    global dbVal
    global procText
    global callText

    Gui, Font, S14 CDefault , Wingdings 3 
    Gui, Add, Button, x210 y3 w40 h30 gResetButton, Q

    Gui, Font, S14 CDefault , Wingdings
    Gui, Add, Button, x255 y3 w40 h30 gQuitButton, x

    Gui, Font, S16 CDefault Bold Italic, Magneto
    Gui, Add, Text, x12 y35 w320 h40 , Telefund Processor v1.0

    ; Gui, Add, Picture, x60 y65 w200 h170 , C:\Users\mmills05\OneDrive - Tufts\AGIS\AutoHotKey\AGIS\Bio\Telefund\Cookie mouth open.png
    Gui, Add, Picture, x60 y65 w200 h170 , Cookie mouth open.png

    Gui, Font, S14 Cblack norm, consolas
    Gui, Add, Text, x22 y240 w130 h30 , Database:
    Gui, Font, S12 Cblack norm, courier new
    Gui, Add, DropDownList, x122 y240 w150 h30 R2 vdbVal, Production||Training 
    
    Gui, Font, S12 Cblack Italic, Times
    Gui, Add, Text, x52 y267 w80 h18 vprocText,
    Gui, Add, Text, x132 y267 w150 h20 vcallText, %callDate% 
   
    ; Generated using SmartGUI Creator 4.0
    Gui, Show, x956 y107 h300 w306, Telefund Processor
    Gui, +AlwaysOnTop
    Return

    ; Generated using SmartGUI Creator 4.0
    ; Gui, Show, x961 y119 h228 w307, Telefund Processor
    ; Return
    QuitButton:
        ExitApp, 

    ResetButton:
        Reload

    GuiDropFiles:
        gui, submit, nohide
        ; msgBox, %dbVal%
        IfEqual, A_GuiEvent, Normal
        {
            return
        }
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
    CommentCRL("")
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
    CommentCRL("")
    return
    
}

MarkFullWithhold(callID)
{   
    ; this can be called whether or not
    ; the records is open

    ensureAdvance()
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
    Send, {TAB 4}
    ;set all fields Y
    Send, Y{TAB}
    Send, Y{TAB}
    Send, Y{TAB}
    Send, Y{TAB}
    Send, Y{TAB}
    Send, Y{TAB}
    Send, Y{TAB}
    Send, Y{TAB}

    ;check for comment
    Send, {TAB 3}
    
    ; CommentCRL("")
    return
    
}
CommentCRL(prepend)
{
    ; need to make this function account
    ; for a filled up comment box
    ; ie, too many chars, advance error window
    global callDate
    ;select All
    Send, {Alt}
    Send, el
    currVal := CopyCell()
    ;modify date format
    StringTrimLeft, dateCalled, callDate, 2
    IfEqual, currVal,
    {
        ;it's empty
        Send, %prepend%Per CRL %dateCalled%
    }Else
    {
        ;it's not empty
        ;does it contain CRL already?
        IfNotInString, currVal, CRL
        {
            Send, {RIGHT}
            Send, `; %prepend%Per CRL %dateCalled%
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
        CommentCRL("")
        FileAppend, 
        (
        `n%A_TAB%Marking ID# %idNum% for No Telefund
        ), %dbVal%_processorLog.txt
        changeMade = True
    } Else IfEqual, callResult, %RemoveList%
    {
        ; MarkNoSolicit(idNum)
        ; TJ decided that RemoveList means Full Withhold
        MarkFullWithhold(idNum)
        CommentCRL("")
        FileAppend, 
        (
        `n%A_TAB%Marking ID# %idNum% for Full Withhold
        ), %dbVal%_processorLog.txt
        changeMade = True
    } Else IfEqual, callResult, Deceased
    {
        MarkFullWithhold(idNum)
        CommentCRL("Deceased ")
        FileAppend, 
        (
        `n%A_TAB%Marking ID# %idNum% for Full Withhold (Deceased)
        ), %dbVal%_processorLog.txt
        changeMade = True
    } Else
    {
        MsgBox, Unrecognized CRL Code (Not Do Not Call, Remove From List, or Deceased.)
        Exit
    }

    ;check for comment
    IfNotEqual, callComment,
    {
        FileAppend, 
        (
        `n%A_TAB%%A_TAB%Comment text: %callComment%
        ), %dbVal%_processorLog.txt
        ;a comment exists
        MsgBox, 1, Comment in CRL, `nComment: `n`n%callComment%
        IfMsgBox, Cancel
        {
            FileAppend, 
            (
            `n%A_TAB%%A_TAB%Previous update may not have been saved.
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
    GuiControl,, procText, processing:
    GuiControl,, callText, %callDate%

    alreadyDone = 0
    Loop, read, %dbVal%_completedCRLs.txt
    {
        IfInString, A_LoopReadLine, %callDate%
        {
            alreadyDone = 1
            MsgBox, The results from %callDate% have already been processed.
            ; break 
            GuiControl,, procText,
            GuiControl,, callText,
            break
        }
    }
    IfEqual, alreadyDone, 1
    {
        return
    }

    Loop, read, %fName%
    {
        
        LineNumber = %A_Index%
        if (LineNumber = 1)
        {
            
            FileAppend,
            (
            `n`nStarting call results from %callDate%...
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
                    ;for now, makr deceased for full withhold
                    updateMailControls(idNum, CallResult, comments)
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
    `n
    Advance IDs processed: %realADVid%
    Other IDs skipped during processing: %notADVid%`n
    ), %dbVal%_processorLog.txt
    FileAppend,
    (
        %callDate%`n
    ), %dbVal%_completedCRLs.txt
    GuiControl,, procText,
    GuiControl,, callText,

    msgBox, Log updated with call results from %callDate%.
}