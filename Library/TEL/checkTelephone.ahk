#Include copyCell.ahk
#Include moveTo.ahk

SetTitleMatchMode, 2

checkTelephone()
{
    ;select area-code field
    Send, {ShiftDown}{CtrlDown}{Right}
    Send, {CtrlUp}{ShiftUp}
    Sleep, 10

    ;copy area code
    areaCode := CopyCell()

    ;Tab to rest of numbers
    Send, {TAB}
    Sleep, 10

    ;select rest of numbers
    Send, {ShiftDown}{CtrlDown}{Right 3}
    Send, {CtrlUp}{ShiftUp}
    Sleep, 25

    ;copy number body
    numberBody := CopyCell()

    wholeNumber = %areaCode%%numberBody%
    fronturl = https://lookups.twilio.com/v1/PhoneNumbers/
    backurl = ?CountryCode=US&Type=carrier&Type=caller-name
    url = %fronturl%%wholeNumber%%backurl%
    ;MsgBox, url equals %url%

    ;Move to firefox and past number into address bar
    MoveTo("Mozilla Firefox")
    Sleep, 250

    ;key to the address bar
    Send, {Alt}
    Sleep, 25
    Send, F
    Sleep, 25
    Send, T
    Sleep, 250

    ;paste address in
    clipboard = %url%
    ClipWait ;
    Send, ^v
    Sleep, 25

    Send, {Enter}

    return
}
