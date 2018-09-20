nextColumn(nextCol = 2, prevCol = 1)
{
    diff := (nextCol - prevCol)
    ;MsgBox, %diff%
    Loop, %diff%
    {
        Send, {TAB}
        Sleep, 50
    }
    return
}