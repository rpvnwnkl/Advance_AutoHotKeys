#Include safeClick.ahk

openPros()
{
    ;this function opens the prospect assignment window
    ;in advance, if it can't it displays a msgbox

    global dbName
    ;click prospect window from bio overview
    Send, {F7}
    safeClick("Button1", dbName, "Custom Tab Profile", "resetGUI", "Prospect Window")
    ;click Prospect INformation button from tracking summary view
    Send, {F7}
    ;need to maintaint window control here
    ;occasionly this click isn't regeistered
    ; safeClick("Button5", dbName, "Prospect Tracking Summary", "resetGUI", "Prospect Tracking Summary Window")
    safeClick("Button1", dbName, "Prospect Tracking Summary", "resetGUI", "Prospect Tracking Summary Window")
    ;click Assignment button from prospect window
    ; this reload seems to help maintain window control
    ; Send, {F7}
    ; safeClick("Button12", dbName, "Prospect - (", "resetGUI", "Prospect Window")
}