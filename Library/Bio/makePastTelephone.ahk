makePastTelephone(currVal)
{
    ;Addtl Business Fax
    IfEqual, currVal, DF
    {
        return "QF"
    }
    ;Additional Business
    IfEqual, currVal, D
    {
        return "Q"
    }
    ;Additional HOme FAx
    IfEqual, currVal, AF
    {
        return "PF"
    }
    ;Addtl Home Phone
    IfEqual, currVal, A
    {
        return "P"
    }
    ;Business Cell
    IfEqual, currVal, BC
    {
        return "PB"
    }
    ;Personal Cell
    IfEqual, currVal, HC
    {
        return "PC"
    }
    ;Seasonal Fax
    IfEqual, currVal, SF
    {
        return "PF"
    }
    ;Home
    IfEqual, currVal, H
    {
        return "P"
    }
    ;Business
    IfEqual, currVal, B
    {
        return "Q"
    }
    ;Home Fax
    IfEqual, currVal, HF
    {
        return "PF"
    }
    ;Business Fax
    IfEqual, currVal, BF
    {
        return "QF"
    }
    ;Seasonal
    IfEqual, currVal, S
    {
        return "P"
    }
    ;Local Phone
    IfEqual, currVal, L
    {
        return "P"
    }
    Exit
}