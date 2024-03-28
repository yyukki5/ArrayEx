Attribute VB_Name = "Sample"
Option Explicit

Sub Sample()
    Dim arr As New ArrayEx2
    Call arr.Init(Range("A1:E3").Value) _
        .DebugPrintAll _
        .Extract("0:1", ":") _
        .DebugPrintAll _
        .GetRow(1) _
        .DebugPrintAll _
        .Skip(1) _
        .SetElement(0, 11) _
        .DebugPrintAll _
        .ToRange(Range("A5"))
End Sub

