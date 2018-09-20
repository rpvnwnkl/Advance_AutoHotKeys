;ID capture subroutine
Capture_ID()
{
    
    ;pull up id window
    Send, {F3}{F3}
    Sleep, 250
    businessID := CopyCell()
    ;close the Id window
    Send, {Esc}
    ;close the entity window
    Send, {CTRLDOWN}{F4}{CTRLUP}
    ;close the entity lookup window
    Send, {CTRLDOWN}{F4}{CTRLUP}
    return businessID
}