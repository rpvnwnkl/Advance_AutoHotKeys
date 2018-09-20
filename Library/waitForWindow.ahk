waitForWindow(searchString, countmax = 50)
{
    ; msgBox, starting string search
    count = 0
    Loop
    {
        IfGreaterOrEqual, count, %countmax%
        {
            return False
        }
        DetectHiddenText, On
        WinGetText, windowCheck 
        ; MsgBox, The text is:`n%windowCheck%
        firstLine = ;
        ;get the first 
        Loop, parse, windowCheck, `n, `r
        {
            ; MsgBox, 4,,%A_LoopField%
            firstLine = %A_LoopField%
            break
        }
        IfInString, firstLine, %searchString%
        {
            ; msgBox, found %searchString%
            return True
        }
        Sleep, 10
        count += 1
    } 
    return True
}