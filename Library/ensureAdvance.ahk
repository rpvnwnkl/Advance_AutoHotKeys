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
    Sleep, 25
}