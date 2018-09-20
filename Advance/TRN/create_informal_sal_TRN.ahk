;Context: You are in the ENT window for an Entity, 
;and you wish to create an informal individual salutation in the SAL window
;Instructions: Press WindowsKey+i 
;The SAL window will stay open afterwards unless you delete the first semi-colon
;in the second to last line of this script
#i::
SetTitleMatchMode, 1
WinWait, ADVTRN, 
IfWinNotActive, ADVTRN, , WinActivate, ADVTRN, 
WinWaitActive, ADVTRN, 
Send, {ALTDOWN}L{ALTUP}{TAB}
;MouseClick, left,  589,  181
;Sleep, 100
Send, {CTRLDOWN}c{CTRLUP}{F5}
WinWait, Go To ..., 
IfWinNotActive, Go To ..., , WinActivate, Go To ..., 
WinWaitActive, Go To ..., 
Send, sal{ENTER}
WinWait, ADVTRN, 
IfWinNotActive, ADVTRN, , WinActivate, ADVTRN, 
WinWaitActive, ADVTRN, 
Send, {F6}{CTRLDOWN}v{CTRLUP}{TAB}{TAB}ii{F8} ;{CTRLDOWN}{F4}{CTRLUP} ;Erase the first semi-colon in this line to have the SAL window close itself
return
