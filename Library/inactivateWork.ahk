inactivateWork()
{
    FormatTime, month,,MM
    FormatTime, day,,dd
    FormatTime, year,,yyyy
    ;##############################            
    ;this either selects id field 
    ;or it shoots past to Job Type or emp name 1
    Send, {TAB}
    ;copy the selected contents
    currCell := CopyCell()
    ;check the copied lengths
    StringLen, empIdLength, currCell
    if empIdLength = 0
    {
        ;no employer id, at field 1
        ;MsgBox, No Employer ID
        Send, {TAB 3}
    }
    if empIdLength > 2
    {
        /*
        this part is necessary because advance highlights fields differently
        depending on whether the window loads with the row selected
        or whether you chose the current row seelction
        */
        ;at field 1 or 2
        ;need to determine which
        StringLeft, firstVal, currCell, 1
        IfEqual, firstVal, 0
        {
            ;MsgBox, Employer ID Present
            Send, {TAB}
        } else
        {
            ;MsgBox, In Field 2
            Send, {Tab 2}
        }
    }
    if empIdLength = 2
    {
        ;employer id, but too far
        ;MsgBox, Emp exists, on Job Type
        ;no movement
    }
    ;now we should be at job type (or therabouts)
    Send, FE
    ;stop day, month, year
    Send, {TAB 12}
    ;Send, %month%
    Send, {TAB}
    ;Send, %day%
    Send, {TAB}
    ;Send, %year%
    ;primary uncheck
    Send, {TAB}
    Send, {SPACE}
    MsgBox, 8192, ,Check the Match Box as it hasn't been updated.
    Send, {F8}
    ;sending enter in case of popups
    Send, {ENTER 2}
    ;f7 to resent cursor to first field
    Send, {F7}
    return
}