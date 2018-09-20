Test_for_Nothing(idFound = True)
{
    ;get window text to check No Entities warning
    Sleep, 1000
    DetectHiddenText, On
    WinGetText, windowCheck 
    ;MsgBox, The text is:`n%windowCheck%
    ;MsgBox 1 idFound = %idFound%

    ;term to check for
    ;windowText = No Entities
    windowText = Looking
    ;Search for windowText in current windows
    IfInString, windowCheck, %windowText%
    {
        idFound := False
        ;Send, {Enter}
        return %idFound%
    }
    return %idFound%
}