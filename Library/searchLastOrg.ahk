searchLastOrg(lastName)
{
    ;open search box, tab to last/org
    Send, {SHIFTDOWN}{F4}{SHIFTUP}
    Send, {TAB 3}
    Send, %lastName%
    Send, {ENTER}
}