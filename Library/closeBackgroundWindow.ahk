#Include ensureAdvance.ahk
#Include closeWindows.ahk

closeBackgroundWindow(num)
{
    BlockInput, On
    ensureAdvance()
    Loop, %num%
    {
        Send, {AltDown}FU
        Send, {AltUp}
    }
    closeWindows(num)
    BlockInput, Off
    return
}