

getProspectID()
{
    global dbName
    #IncludeAgain, CopyCell.ahk
    
    delay = 250
    ; this gives 'control' to the current window
    ; Send, {ENTER}
    ; Send, {CtrlDown}{F4}{CtrlUp}
    ; alt method below
    Sleep, %delay%
    Send, {F7}
    Send, {ShiftDown}{TAB 4}
    Send, {ShiftUp}
    ;now we can tab to prospect window
    ;on the main bio view
    ;Send, {TAB 8}
    Send, {Enter}
    Sleep, %delay%

    ;now pop up a dialog and exit
    ;allowing us to navigate this window
    Send, {F3}
    Send, {ESC}
    Sleep, %delay%
    ;now tab to assignment box
    Send, {TAB 3}
    Send, {Enter}
    Sleep, %delay%
    ;now would be a good time to capture
    ;prospect ID, if you're interested
    Send, {F3}
    prosID := CopyCell()
    ;close the windows
    Send, {ESC}
    Send, {ShiftDown}{F4 2}{ShiftUp}
    ;now nothing should be open
    return prosID
    
}