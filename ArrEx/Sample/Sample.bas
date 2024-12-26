Attribute VB_Name = "Sample"
'<dir .\Sample /dir>
Option Explicit

Sub Sample()

    Dim ar, rearr
    ar = array2d

    Call ArrEx(array2d) _
        .DebugPrint("," & vbTab) _
        .WhereBy(1, aexGreaterThan, 2) _
        .DebugPrint("," & vbTab) _
        .OrderByDescending(1) _
        .DebugPrint("," & vbTab) _
        .SelectRows(1, 2) _
        .DebugPrint("," & vbTab) _
        .SelectColumns(2, 3) _
        .ConvertTo(vbString) _
        .DebugPrint("," & vbTab)

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
