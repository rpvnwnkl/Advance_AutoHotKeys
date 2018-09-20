
SetTitleMatchMode, 2
;set location of batch file
folder_location = t:\interfaces and shadow systems\salesforce\prod\
batch_filename = tu_load_sfdc_olc_run_prod.bat

Run, %batch_filename%, %folder_location%

WinWait, cmd.exe, 
IfWinNotActive, cmd.exe, , WinActivate, cmd.exe, 
WinWaitActive, cmd.exe,

Sleep, 100

Loop, Read, %A_MyDocuments%\DB_Password\logpass.txt
{
    Loop, parse, A_LoopReadLine, `n,`r
    {
        IfEqual, A_Index, 1
        {
            ;First line is usrNme
            Send, %A_LoopField%
            Send, {Enter}
        } 
        IfEqual, A_Index, 2
        {
            ;Second line is pWrod
            Send, %A_LoopField%
            Send, {Enter}
        }
    }
}

return

