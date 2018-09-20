;Context: You are in the ENT window for an Entity, 
;and you wish to create an informal individual salutation in the SAL window
;Instructions: Press WindowsKey+i 
;The SAL window will stay open afterwards unless you delete the first semi-colon
;in the second to last line of this script

#Include createInformal.ahk
#Include copyCell.ahk
#Include windowPresent.ahk

newEntitySalutation()
{
    global dbName
    ;SetTitleMatchMode, 2
    If WinActive(dbName)
    {
         ;######################
         ;MUST BE IN ENT WINDOW
         ;########################
        If windowPresent("Entity - (ID")
        {
            ;Go to "Last" name category, TAB to first
            Send, {ALTDOWN}L{ALTUP}{TAB}
            ;copy first name
            firstName := copyCell()
            createInformal(firstName)
        }

    }
    return
}
