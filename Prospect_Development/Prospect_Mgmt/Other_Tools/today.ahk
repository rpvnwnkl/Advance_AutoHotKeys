today()
{
    FormatTime, month,,MM
    FormatTime, day,,dd
    FormatTime, year,,yyyy
    today = %month%/%day%/%year%
    return today
}

#IfWinActive, Production Database
^LButton::
today := today()
Send, %today%