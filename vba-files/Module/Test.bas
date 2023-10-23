Attribute VB_Name = "Test"
Option Explicit

Private Function array2d() As Variant
Dim a
ReDim a(1 To 5, 1 To 10)
Dim i, j
For i = 1 To 5
    For j = 1 To 10
        a(i, j) = i * j
    Next j
Next i
array2d = a
End Function

Sub Init_NoException()
Dim a2d As New ArrayEx2
a2d.Init (array2d())

Dim b2d As New ArrayEx2
Dim c2d As New ArrayEx2
b2d.Init (a2d)
c2d.Init (a2d())

abc (a2d.Mine)
abc (a2d.Mine)

End Sub

Private Function abc(ByRef a2d As Object)
Debug.Print a2d.Value(1, 1)
End Function


Sub GetValue_NoEception()
Dim a2d As New ArrayEx2
a2d.Init (array2d())

Debug.Assert IsNull(a2d.Value) = False
Debug.Assert a2d.Value(1, 2) = 1 * 2
Debug.Assert UBound(a2d.Value(":", 1)) = 5
Debug.Assert UBound(a2d.Value(2, ":")) = 10
Debug.Assert UBound(a2d.Value, 1) = 5
Debug.Assert UBound(a2d.Value, 2) = 10
Debug.Assert UBound(a2d.Value(rows:=2)) = 10
Debug.Assert UBound(a2d.Value(cols:=1)) = 5
End Sub
