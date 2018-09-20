; Example #4: Parse a comma separated value (CSV) file:
Loop, read, C:\Users\mmills05\Documents\AGIS\Data\Degrees\SIS_Ids.csv 
{
    LineNumber = %A_Index%
    Loop, parse, A_LoopReadLine, CSV 
    {
        MsgBox, 4, , Field %LineNumber%-%A_Index% is:`n%A_LoopField%`n`nContinue?
        IfMsgBox, No
        return    
    }
}