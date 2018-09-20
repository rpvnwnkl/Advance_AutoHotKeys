Gui, Font, s10, Courier New
Gui, Add, Edit, w640 r33 vMyEdit , < Drop Your Text Files Here >
Gui, Show
Return

GuiDropFiles:
    StringSplit, F, A_GuiEvent, `n
    Loop, %F0%
    {
        currVal := F%A_Index%
        ; currVal = %currVal%
        MsgBox, %currVal%
    }
;  FileRead, Text, %F1%
;  GuiControl,, MyEdit, %F1%
    Return

GuiClose:
GuiEscape:
 ExitApp