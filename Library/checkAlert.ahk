#Include windowPresent.ahk
#Include closeWindows.ahk

checkAlert(delay = 150)
{
    ; check for alert
    Sleep, %delay%
    alertFound := windowPresent("Alerts -")
    ; msgBox, %alertFound%
    IfEqual, alertFound, 1
    {
        closeWindows()
    }
}