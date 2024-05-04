Attribute VB_Name = "Sample"
Option Explicit

Sub Sample()

    Dim ar, rearr
    ar = array2d
    rearr = ArrayExCore.HSTACK(ar, ar)

    Dim arr As New ArrayEx2
    Call arr.Init(rearr) _
        .DebugPrintAll _
        .Extract("1:3", ":") _
        .WhereEvaluated("x", 1, "x>1") _
        .DebugPrintAll _
        .GetRow(1) _
        .DebugPrintAll _
        .OrderByAscending _
        .SetElement(1, 11) _
        .DebugPrintAll

End Sub

Private Function array2d()
    Dim a, i, j
    ReDim a(1 To 5, 1 To 10)
    For i = 1 To 5
        For j = 1 To 10
            a(i, j) = i * j
        Next j
    Next i
    array2d = a
End Function
