
;Pressing windows key + A reloads the script in case it gets wonky

#Include moveTo.ahk

SetTitleMatchMode, 2

saveTeamRequest()
{
    ;make sure we are in Team
    moveTo("Alexsys Team 2 Pro")

    ;Ctrl+W opens team request from list view
    Send, {CTRLDOWN}w{CTRLUP}
    Sleep, 300
    ;make sure ticket window is open
    moveTo("Team Work Request") 
    Sleep, 300

    ;requird that no fields are already selected
    ;F9 sends 'edit' command
    Send, {F9}
    Sleep, 25
    Send, {TAB}{TAB}{TAB}

    ;start work in progress
    Send, w
    Send, {TAB}
    Sleep, 50

    ;save changes with Ctrl+S
    Send, {CTRLDOWN}s{CTRLUP}
    Sleep, 300
    ;Close with Alt+F4
    Send, {ALTDOWN}{F4}{ALTUP}
    return
}