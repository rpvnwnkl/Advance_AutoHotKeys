CopyCell()
{
    Sleep, 10
    oldClipboard = %clipboard%
    clipboard =
    Send, {CTRLDOWN}c{CTRLUP}
    ClipWait, 1 ;
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
    clipboard = %oldClipboard%
    return cellOfInterest
}