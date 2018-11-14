createInformal(informalName)
{
    ;call up salutation window
    callVar = SAL
    callGoTo(callVar) 
    ;Now create a new entry
    Send, {F6}
    ;send informal salutation
    Send, %informalName%
    ;move to type field
    Send, {TAB}{TAB}
    ;set type as informal individual
    Send, ii
    ;save and then exit
    Send, {F8}
    Send, {CTRLDOWN}{F4}
    Send, {CTRLUP} 
}