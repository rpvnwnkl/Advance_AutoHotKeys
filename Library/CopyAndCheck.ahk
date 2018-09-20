CopyAndCheck()
{
    ;Copy highlighted ID
    ;global
    entityID := CopyCell()

    IfEqual, entityID, 
    {
        ;MsgBox, Done
        return
    }

    IfWinNotExist, Production Database
    {
        MsgBox, Advance isn't open
        return
    }

    ;move to Advance
    MoveTo("Production Database")
    Sleep, 100 
    return entityID   
}