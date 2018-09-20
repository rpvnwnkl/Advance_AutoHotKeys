searchAgain()
{
    MsgBox, 262145,, Record Found?            
    
        IfMsgBox Cancel 
        {
            MsgBox, Search Again?
        }
        ;check if profile was opened or not
        WinGetText, windowCheck
        ;Search for windowText in current windows
        windowText = Custom
        IfNotInString, windowCheck, %windowText%
        {
            Send, {Enter} 
        }
    }
    return idFound
}