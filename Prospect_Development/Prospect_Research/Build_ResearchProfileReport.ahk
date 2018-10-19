; Context: You want to generate a Research profile report
; First, open Advance to the record for which you'd like a report created
; Start script with WIN + s
; Note: Advance must be the active window or else script will not start
#IfWinActive Production Database
#s::
SetTitleMatchMode, 2
Sleep, 100
SendInput, {ALTDOWN}fo{ALTUP}
Sleep, 500
SendInput, {DOWN 2}
Sleep, 200
SendInput, {ENTER}
Sleep, 200
SendInput, {TAB 2}
Sleep, 100
SendInput, {DOWN 18}
Sleep, 100
SendInput, {ENTER}
Sleep, 250
SendInput, {TAB}{SPACE}
Sleep, 100
SendInput, {TAB}{SPACE}
Sleep, 100
SendInput, {TAB}{SPACE}
Sleep, 100
SendInput, {TAB 2}
Sleep, 100
SendInput, {TAB}{SPACE}
Sleep, 100
SendInput, {TAB}
Sleep, 100
SendInput, {TAB}{SPACE}
Sleep, 100
SendInput, {TAB}{SPACE}
Sleep, 100
SendInput, {TAB}{SPACE}
Sleep, 100
SendInput, {TAB}{SPACE}
Sleep, 100
SendInput, {ENTER}
Sleep, 2500
SendInput, {ALTDOWN}fd{ALTUP}
Sleep, 100
SendInput, {ENTER}
Sleep, 2000
Click, 437, 724
Sleep, 1000
SendInput, {ENTER}
Sleep, 500
if WinExist("ahk_class #32770") {
    SendInput, {LEFT}
    Sleep, 100
    SendInput, {ENTER}
}
Return