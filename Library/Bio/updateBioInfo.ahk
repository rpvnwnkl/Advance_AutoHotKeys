
#Include, windowPresent.ahk
#Include, copyCell.ahk
#Include, Bio\makePastAddress.ahk
#Include, Bio\makePastTelephone.ahk
#Include, waitForWindow.ahk

updateEmail(newStatus)
{
    currSelection := CopyCell()
    IfEqual, currSelection,
    {
        Send, {TAB}
    }
    IfEqual, newStatus, K
    {
        Send, {TAB 2}
        sleep, 10
        Send, K
        sleep, 10
        Send, {F8}
    } Else IfEqual, newStatus, P
    {
        Send, {TAB}
        ; Send, {Space}
        Send, {TAB}
        Send, I
        Send, {TAB 6}
        Send, BIO
        Send, {F8}
    }
    return
}

updateAddress(newStatus)
{
    IfEqual, newStatus, K
    {
        Send, {ENTER}
        Sleep, 25
        Send, {AltDown}Y{AltUp}
        Send, {TAB}
        Send, K
        Send, {F8}
    } Else IfEqual, newStatus, P
    {
        Send, {ENTER}

        waitForWindow("Address - ")
        Send, {AltDown}Y
        SEnd, {AltUp}
        currVal := copyCell()
        ; msgbox, %currVal%
        newVal := makePastAddress(currVal)
        Sleep, 10
        Send, %newVal%
        Send, {TAB}
        Send, I
        Send, {TAB 7}
        ; Send, {Space}
        Send, {F8}
        ;in case it got marked preferred
        ; IfWinActive, Address
        ; {
        ;     Send, {ENTER}
        ;     Send, {AltDown}Y{AltUp}
        ;     send, {Tab 8}
        ;     Send, {Space}
        ;     Send, {F8}
        ; }

    }
    return
}

updateTelephone(newStatus)
{
    currSelection := CopyCell()
    IfEqual, currSelection,
    {
        Send, {TAB}
    }
    IfEqual, newStatus, K
    {
        Send, {TAB 6}
        Send, K
        Send, {F8}
    } Else IfEqual, newStatus, P
    {
        currVal := CopyCell()
        newVal := makePastTelephone(currVal)
        Send, %newVal%
        Send, {TAB}
        ; Send, {Space}
        Send, {Tab 5}
        Send, I
        Send, {F8}
    }
    return
}

updateEcon(newStatus)
{
    currSelection := CopyCell()
    IfEqual, currSelection,
    {
        Send {Tab 3}
    } Else
    {
        Send, {TAB 2}
    }
    IfEqual, newStatus, K
    {
        Send, K
    } Else IfEqual, newStatus, P
    {
        Send, I
        Send, {tab 8}
        ; Send, {Space}
    }
    SEnd, {F8}
    return
}
updateEmployment(newStatus)
{
    Send, {TAB 3}
    Send, FE
    Send, {TAB 15}
    Send, {Space}

    SEnd, {F8}
    return
}
#Include, windowPresent.ahk
updateBioInfo(updateVal)
{
    Sleep, 10
    ;mark as last known
    moveTo(dbName)
    If windowPresent("Email")
    {
        updateEmail(updateVal)
    } 
    Else If windowPresent("Addresses")
    {
        updateAddress(updateVal)
    } 
    Else If windowPresent("Telephone")
    {
        updateTelephone(updateVal)
    } 
    Else If windowPresent("eContact")
    {
        updateEcon(updateVal)
    }
    Else If windowPresent("Employment")
    {
        updateEmployment(updateVal)
    }
    return
}