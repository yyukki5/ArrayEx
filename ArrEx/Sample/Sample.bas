Attribute VB_Name = "Sample"
'<dir .\Sample /dir>
Option Explicit

Sub Sample()

    Call ArrEx(array2d) _
        .DebugPrint("," & vbTab) _
        .RedimPreserve(aexRank2, 1, 4, 1, 5) _
        .DebugPrint("," & vbTab) _
        .WhereBy(2, aexGreaterThan, 2) _
        .DebugPrint("," & vbTab) _
        .OrderByDescending(1) _
        .DebugPrint("," & vbTab) _
        .SelectRows(1, 2) _
        .DebugPrint("," & vbTab) _
        .SelectColumns(2, 3) _
        .DebugPrint("," & vbTab) _
        .LeftJoin(HelloWorlds, 1, 1) _
        .DebugPrint("," & vbTab) _
        .SelectColumns(4) _
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

Private Function HelloWorlds()
    Dim a, i As Long
    ReDim a(1 To 20, 1 To 2)
    For i = 1 To 20
        a(i, 1) = i
        a(i, 2) = "Hello World - " & i
    Next i
    HelloWorlds = a
End Function

