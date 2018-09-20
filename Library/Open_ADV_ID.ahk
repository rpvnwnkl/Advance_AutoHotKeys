#Include, CopyCell.ahk
#Include, moveTo.ahk
#Include, pullUpEntity.ahk
; SetTitleMatchMode, 2
Open_ADV_ID()
{
    global dbName
    
    ogClipContents = %clipboard%

    ;copy entity ID
    entityID := CopyCell()

    IfEqual, entityID, 
    {
        ;MsgBox, Done
        return
    }

    IfWinNotExist, %dbName%
    {
        BlockInput, Off
        MsgBox, Advance isn't open
        return
    } else 
    {
        moveTo(dbName)
        pullUpEntity(entityID)
    }

    clipboard = %ogClipContents%
    return
}