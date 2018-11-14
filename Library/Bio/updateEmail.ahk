updateEmail(newStatus)
{
    Send, {TAB 2}
    sleep, 10
    Send, %newStatus%
    sleep, 10
    Send, {F8}
}