Attribute VB_Name = "Sample"
Option Explicit

Sub Sample()

    Dim ar, rearr
    ar = Range("A1:E3").Value
    rearr = ArrayExCore.HSTACK(ar, ar)

    Dim arr As ArrayEx2
    Set arr = New ArrayEx2
    Call arr.Init(Range("A1:E3").Value) _
        .DebugPrintAll _
        .Extract("1:3", ":") _
        .DebugPrintAll _
        .GetRow(1) _
        .DebugPrintAll _
        .Skip(1) _
        .SetElement(1, 11) _
        .DebugPrintAll _
        .ToRange(Range("A5"))
    
End Sub
