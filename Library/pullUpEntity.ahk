#Include ensureAdvance.ahk

pullupEntity(entID)
{
    BlockInput, On
    ensureAdvance()

    ;pull up entity window
    Send, {ShiftDown}{F4}{ShiftUp}
    Sleep, 25
    Send, %entID%
    Send, {Enter}
    BlockInput, Off
}