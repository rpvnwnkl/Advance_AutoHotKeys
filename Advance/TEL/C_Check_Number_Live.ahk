
;Context: You are in the TEL window and want to check whether a number \
;           is a cell or a landline
;	 
;	
;Instructions: have advance open with cursor in the area code field and
;              
;		Press windows key and the letter C
;       Pressing Windows Key and Letter A will reset in case of error


CopyCell()
{
    Sleep, 10
    clipboard = ;
    Send, {CTRLDOWN}c{CTRLUP}
    ClipWait, 0 ;
    cellOfInterest = %clipboard%
    ;clean up id from excel
    ;check if you are in excel so the newline can be dropped
    IfWinActive, Excel, , , 
    {
        ;clean up id from excel
        StringTrimRight, clipboard, clipboard, 2
        Sleep, 10
        cellOfInterest = %clipboard%
        Send, {Esc}
    }
    clipboard = ;
    return cellOfInterest
}
MoveTo(destination)
{
    ;Move to destination 
    WinWait, %destination%
    IfWinNotActive, %destination%, 
    ;MsgBox "2"
    ;#WinActivateForce
    WinActivate, %destination%, ,
    ;MsgBox "4"
    ;WinWaitActive, %destination%
    WinWaitActive, %destination%, 

}

SetTitleMatchMode, 2
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
;Send, %url%

;Msgbox OK?

Send, {Enter}

return
