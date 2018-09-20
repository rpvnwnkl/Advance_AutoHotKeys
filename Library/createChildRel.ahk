createChildRel(childID)
{
    ;call up child window
    callVar = CHILD
    callGoTo(callVar)
    ;Now create a new entry
    Send, {F6}CH{TAB}
    Send, %childID%
    ;save and then exit
    Send, {F8}
    Send, {CTRLDOWN}{F4}{CTRLUP}
    
}