inactivateAddress(comment)
{
    global ;
    ;go to type field
    Send, {AltDown}Y{AltUp}
    addrType := CopyCell()
    If addrType in A,H
    {
        Send, P
    } Else If addrType in B,D
    {
        Send, Q
        addrLinked = False
        MsgBox, 4, EMPT Info, Is the Business Address Linked to an employment record?
        IfMsgBox, Yes
        {
            addrLinked = True
        }

    } Else If addrType in S
    {
        Send, T
    } Else If addrType in C,L
    {
        Send, U
    } Else
    {
        MsgBox, Unknown Address Type
        return
    }

    ;go to status field
    Send, {TAB}
    addrStatus := CopyCell()
    If addrStatus in I,N,K
    {
        ;changes have been made and no more are necessary
        MsgBox, This address has already been updated
        return
    } Else
    {
        ;mark inactive
        Send, I
    }

    /*
    ;go to preferred box
    Send, {TAB 7}
    Send, {Space}
    */
    ;capture start month
    Send, {TAB}
    Send, {ShiftDown}{Right 2}{ShiftUp}
    modM := CopyCell()

    ;capture start day
    Send, {TAB}
    Send, {ShiftDown}{Right 2}{ShiftUp}
    modD := CopyCell()

    ;capture start month
    Send, {TAB}
    Send, {ShiftDown}{Right 4}{ShiftUp}
    modY := CopyCell()


    If addrLinked
    {
        ;skip linked box
        Send, {Tab 5}
    } Else
    {
        ;go to preferred box
        Send, {TAB 4}
    }
    ;uncheck preferred
    Send, {Space}
    
    ;capture source code
    Send, {TAB}
    modCode := CopyCell()

    ;go to comment field
    Send, {TAB 3}
    ;select comment field and save to variable
    modComment := CopyCell()
    IfEqual, modComment, 
    {
        Send, %comment%
    } Else
    {
        Send, {Right}; %comment%
    }

    ;now save, return
    Send, {F8}

    ;for telephone record
    Send, {ENTER}
    ;for business info
    if addrLinked
    {
        Send, {ENTER}
    }
    ;close two addr windows
    Sleep, 100
    Send, {CtrlDown}{F4}{CtrlUp}
    Sleep, 100
    Send, {CtrlDown}{F4}{CtrlUp}

}