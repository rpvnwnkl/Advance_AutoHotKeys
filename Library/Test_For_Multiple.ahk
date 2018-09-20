Test_for_Multiple(idFound)
{
    ;get window text to check for multiple hits
    WinGetText, windowCheck 
    ;MsgBox, The text is:`n%windowCheck%

    ;term to check for
    windowText = Criteria
    ;Search for windowText in current windows
    IfInString, windowCheck, %windowText%
    {
        ;WinSetTitle, SIS Converter, , NewTitle
        ControlSetText, Static1, User Input Required..., SIS Converter
        MsgBox, 262145,, Please open the correct entity record            
    
        IfMsgBox Cancel 
        {
            Exit
            idFound = False
            return idFound
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