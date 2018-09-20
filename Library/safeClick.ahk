#Include waitForWindow.ahk
#Include moveTo.ahk

safeClick(buttonName, windowName, windowText, callBack, easyWindowName)
{
    ;this function opens windows in advance, using mouse buttons
    ;be careful when using a callback.

    ;buttonName, windowName, windowText are all as described in ControlClick docs
    ;easyWindowName is used for display message upon failure
    ;callback can execute custom behavior on failure (ie, it calls a function)
    ;set callback to 0 to default as return

    moveTo(windowName)
    delay = 150
    windowFound := waitForWindow(windowText, delay)
    IfEqual, windowFound, 1
    {
        Sleep, %delay%
        ControlClick, %buttonName%, %windowName%, %windowText%,,2,NA
    } Else
    {
        BlockInput, Off
        Msgbox, Could not open %easyWindowName%
        IfNotEqual, callBack, 0
        {
            %callBack%()
        } Else
        {
            return
        }
        
    }
}