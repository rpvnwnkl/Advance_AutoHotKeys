marry(spouse1, spouse2, fatherFirst, fatherMiddle, fatherLast, fatherSuffix, motherFirst, motherMiddle, motherLast)
{
    ;pull up spouse1 record
    Send, {ShiftDown}{F4}{ShiftUp}
    Sleep, 50
    Send, %spouse1%
    Send, {ENTER}

    ;call up marriage window
    callVar = Mar
    callGoTo(callVar)
    Sleep, 50

    ;status is Married - M
    Send, M{TAB}
    ;spouse ID
    Send, %spouse2%{TAB}
    
    ;enter joint names
    If fatherLast <> %motherLast% 
    {
        ;we are stacking names
        
        ;find middle initial
        StringTrimRight, fatherInitial, fatherMiddle, StrLen(fatherMiddle)-1
        StringTrimRight, motherInitial, motherMiddle, StrLen(motherMiddle)-1
        
        ;enter appropriately given whether initial exists
        If StrLen(fatherInitial) > 0
        {
            Send, Mr. %fatherFirst% %fatherInitial%. %fatherLast%
        } else
        {
            Send, Mr. %fatherFirst% %fatherLast%
        }
        ;Father suffix (jr, etc)
        If fatherSuffix is not space
        {
            Send, , %fatherSuffix%
            Send, .
        }
        ;now mother
        Send, {TAB}
        If StrLen(motherInitial) > 0
        {
            Send, Ms. %motherFirst% %motherInitial%. %motherLast%
        } else
        {
            Send, Ms. %motherFirst% %motherLast%
        }
        
        ;build jnt sal now
        Send, {TAB}
        Send, Mr. %fatherLast% and Ms. %motherLast%
        Send, {Tab 13}
    } else 
        {
            ;we accept default naming scheme
            Send, {TAB 15}
        }
    Send, Team 240749
    ;save and then uncheck home phone
    Send, {F8}
    Sleep, 50
    Send, {Enter}
    Send, {Space}
    Send, {F8}

    ;add joint informal salutation in sal window
    callVar = SAL
    callGoTo(callVar)
    Send, {F6}
    Send, %fatherFirst% and %motherFirst%
    Send, {Tab 2}ji{F8}
    Send, {SHIFTDOWN}{F3}{SHIFTUP}
    Send, {F6}
    Send, %fatherFirst% and %motherFirst%
    Send, {Tab 2}ji{F8}
    Send, {CTRLDOWN}{F4}{CTRLUP}

    ;close marriage window
    Send, {CTRLDOWN}{F4}{CTRLUP}
    ;close entity window
    Send, {CTRLDOWN}{F4}{CTRLUP}
    return

}