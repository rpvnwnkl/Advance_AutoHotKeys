
#Include, windowPresent.ahk
#Include, copyCell.ahk

; updateEmail(newStatus)
; {
;     IfEqual, newStatus, K
;     {
;         Send, {TAB 2}
;         sleep, 10
;         Send, K
;         sleep, 10
;         Send, {F8}
;     } Else IfEqual, newStatus, P
;     {
;         Send, {TAB}
;         Send, {Space}
;         Send, {TAB}
;         Send, I
;         Send, {TAB 6}
;         Send, BIO
;         Send, {F8}
;     }
;     return
; }

; updateAddress(newStatus)
; {
;     IfEqual, newStatus, K
;     {
;         Send, {ENTER}
;         Sleep, 25
;         Send, {AltDown}Y{AltUp}
;         Send, {TAB}
;         Send, K
;         Send, {F8}
;     } Else IfEqual, newStatus, P
;     {
;         Send, {ENTER}
;         Sleep, 25
;         Send, {AltDown}Y{AltUp}
;         currVal := copyCell()
;         newVal := makePastAddress(currVal)
;         Send, %newVal%
;         Send, {TAB}
;         Send, I
;         Send, {TAB 7}
;         Send, {Space}
;         Send, {F8}
;         ;in case it got marked preferred
;         IfWinActive, Address
;         {
;             Send, {ENTER}
;             Send, {AltDown}Y{AltUp}
;             send, {Tab 8}
;             Send, {Space}
;             Send, {F8}
;         }

;     }
;     return
; }

; updateTelephone(newStatus)
; {
;     currSelection := CopyCell()
;     IfEqual, currSelection,
;     {
;         Send, {TAB}
;     }
;     IfEqual, newStatus, K
;     {
;         Send, {TAB 6}
;         Send, K
;         Send, {F8}
;     } Else IfEqual, newStatus, P
;     {
;         currVal := CopyCell()
;         newVal := makePastTelephone(currVal)
;         Send, %newVal%
;         Send, {TAB}
;         Send, {Space}
;         Send, {Tab 5}
;         Send, I
;         Send, {F8}
;     }
;     return
; }

; updateEcon(newStatus)
; {
;     currSelection := CopyCell()
;     IfEqual, currSelection,
;     {
;         Send {Tab 3}
;     } Else
;     {
;         Send, {TAB 2}
;     }
;     SEnd, K
;     SEnd, {F8}
;     return
; }
updateCTSInfo(updateKey)
{
    global dbName
    ;keys are sticking
    Send, {CtrlUp}
    1701 := "Unbooked; meets age/reunion BW booking criteria"
    1702 := "Unbooked; not meet BW criteria; have docs for future"
    1703 := "Unbooked; meets BW criteria, but won't share $$/docs"
    1704 := "Unbooked; does not meet age/reunion BW criteria"
    1705 := "Ubooked; meets BW criteria, but gift is contingent"
    1706 := "Prospect has a booked planned gift"

    Sleep, 10
    MoveTo(dbName)
    Click, 250, 300
    Send, {TAB 4}
    Send, {ShiftDown}{End}
    SEnd, {ShiftUp}
    ; msgBox, %updateKey%
    Send, %updateKey%
    Send, {Tab 6}
    Send, BIO
    Send, {TAB}
    Send, {ShiftDown}{End}
    SEnd, {ShiftUp}
    keyVal := %updateKey%
    ; msgbox, %keyVal%
    Send, %keyVal%
    Send, {F8}
    ; ControlGet, childWin, Hwnd,,FNWND31202,, 
    ; ControlClick, , ahk_id %childWin% 
    return
}