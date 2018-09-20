;returns true/false

windowPresent(windowName)
{
    DetectHiddenText, On
    WinGetText, windowText 
    ; MsgBox, The text is:`n%windowText%
    firstLine = ;
    ;get the first 
    Loop, parse, windowText, `n, `r
    {
        ; MsgBox, 4,,%A_LoopField%
        firstLine = %A_LoopField%
        break
    }
    IfInString, firstLine, %windowName%
    {
        ; msgBox, found %searchString%
        return True
    }
    return False
}