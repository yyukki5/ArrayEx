Attribute VB_Name = "Sample"
'<dir .\Sample /dir>
Option Explicit

Sub Sample()

    Call ArrayEx(array2d) _
        .DebugPrint("," & vbTab) _
        .WhereBy("x => x(1) > 2") _
        .DebugPrint("," & vbTab) _
        .SelectBy("x=>x(2) + x(3)") _
        .DebugPrint("," & vbTab) _
        .OrderByDescending _
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
