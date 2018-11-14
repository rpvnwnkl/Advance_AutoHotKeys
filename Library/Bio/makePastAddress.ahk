makePastAddress(currVal)
{
    ;Home
    IfEqual, currVal, H
    {
        return "P"
    }
    ;Additional Home
    IfEqual, currVal, A
    {
        return "P"
    }
    ;Seasonal
    IfEqual, currVal, S
    {
        return "T"
    }
    ;Business
    IfEqual, currVal, B
    {
        return "Q"
    }
    ;Additional Business
    IfEqual, currVal, D
    {
        return "Q"
    }
    ;Campus
    IfEqual, currVal, D
    {
        return "U"
    }
    ;local
    IfEqual, currVal, L
    {
        return "U"
    }
    return currVal
}