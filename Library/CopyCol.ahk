CopyCol()
{
    Sleep, 10
    clipboard = ;
    Send, {ShiftDown}{CtrlDown}{Down}{ShiftUp}{CtrlUp}
    Send, {CTRLDOWN}c{CTRLUP}
    ClipWait, 0 ;
    cellOfInterest = %clipboard%
    /*
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
    */
    clipboard = ;
    Send, {ESC}
    return cellOfInterest
}