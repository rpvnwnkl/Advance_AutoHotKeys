getToProsA()
{
    ; this gives 'control' to the current window
    Send, {ENTER}
    Send, {CtrlDown}{F4}{CtrlUp}

    ;now we can tab to prospect window
    ;on the main bio view
    Send, {TAB 8}
    Send, {Enter}

    ;now pop up a dialog and exit
    ;allowing us to navigate this window
    Send, {F3}
    Send, {ESC}

    ;now tab to assignment box
    Send, {TAB 3}
    Send, {Enter}

    ;now would be a good time to capture
    ;prospect ID, if you're interested

    ;open Assignment Window
    Send, {AltDown}A{AltUp}
    
    ;move control back to advance, jik
    MoveTo("Production Database") 
    
    
    return
}