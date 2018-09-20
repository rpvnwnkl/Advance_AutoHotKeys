#Include ensureAdvance.ahk

closeWindows(numWindows = 1)
{
    ;close windows
    ; this is the old version, doesn't seem to work when
    ; blockinput is on
    ; Send, {CtrlDown}{F4 %numWindows%}{CtrlUp}

    BlockInput, On
    ensureAdvance()
    ;new version
    Loop, %numWindows%
    {
        Send, {AltDown}FC
        Send, {AltUp}
        Sleep, 5
    }
    BlockInput, Off
    return
    
}