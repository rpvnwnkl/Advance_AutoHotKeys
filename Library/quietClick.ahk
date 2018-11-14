#Include waitForWindow.ahk
#Include moveTo.ahk

quietClick(buttonName, windowName, windowText, callBack)
{
    ;this function opens windows in advance, using mouse buttons
    ;be careful when using a callback.

    ;buttonName, windowName, windowText are all as described in ControlClick docs
    ;callback can execute custom behavior on failure (ie, it calls a function)
    ;set callback to 0 to default as return

    moveTo(windowName)
    delay = 150
    windowFound := waitForWindow(windowText, delay)
    IfEqual, windowFound, 1
    {
        Sleep, %delay%
        ControlClick, %buttonName%, %windowName%, %windowText%,,2,NA
        return
    } Else
    {
        IfNotEqual, callBack, 0
        {
            %callBack%()
        } Else
        {
            return
        }
        
    }
}