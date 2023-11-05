Attribute VB_Name = "Sample"
Option Explicit

Sub Sample()

Dim arr As New ArrayEx2
Call arr.Init(Range("A1:E3").Value) _
    .DebugPrintAll _
    .Extract("1:2", "3:5") _
    .DebugPrintAll _
    .GetRow(2) _
    .DebugPrintAll _
    .ToRange(Range("A5"))
    
End Sub
