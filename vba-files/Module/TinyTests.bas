Attribute VB_Name = "TinyTests"
Option Explicit

Sub RunTests()
    ' before using, switch error check in Tools > Options > General > Error trap
    Debug.Print "--- Start tests (" & Format(Now) & ") ---"
    
    RunTest "Init_NoException"
    RunTest "Create_NoException"
    RunTest "GetValue_NoEception"
    RunTest "SetValue_GetMethodIsCorrect_NoException"
    RunTest "AddElement_NoException"
    RunTest "Converts_NoException"
    '    RunTest "DebugPrint_NoException"
    RunTest "WrapFunctions_NoException"
    RunTest "Array1Linqs_NoException"
    RunTest "Array2Linqs_NoException"
    
    Debug.Print "--- Finish tests ---"
End Sub

Private Function array1d() As Variant
    Dim a
    ReDim a(1 To 5)
    Dim i
    For i = 1 To 5
        a(i) = i
    Next i
    array1d = a
End Function

Private Function array1drow() As Variant
    Dim a
    ReDim a(1 To 10)
    Dim i
    For i = 1 To 10
        a(i) = i
    Next i
    array1drow = a
End Function

Private Function array1dHasDuplication() As Variant
    Dim a:      a = array1d:      a(3) = 2
    array1dHasDuplication = a
End Function

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

Private Function array2dDuplication() As Variant
    Dim a: a = array2d
    a(3, 2) = 4
    array2dDuplication = a
End Function

Sub Init_NoException()
    On Error Resume Next
    
    Dim a0d As New ArrayEx0
    a0d.Init (1)
    If AssertHasError Then Exit Sub
    a0d.Init (array1d)
    If AssertHasNoError Then Exit Sub
    a0d.Init (array2d)
    If AssertHasNoError Then Exit Sub
    
    Dim a1d As New ArrayEx1
    a1d.Init (1)
    If AssertHasNoError Then Exit Sub
    a1d.Init (array1d)
    If AssertHasError Then Exit Sub
    a1d.Init (array2d)
    If AssertHasNoError Then Exit Sub
    
    Dim a2d As New ArrayEx2
    a2d.Init (1)
    If AssertHasNoError Then Exit Sub
    a2d.Init (array1d())
    If AssertHasNoError Then Exit Sub
    a2d.Init (array2d())
    If AssertHasError Then Exit Sub
    
    If AssertTrue(a2d.Equals(array2d())) Then Else Exit Sub

End Sub

Sub Create_NoException()
    On Error Resume Next
    
    Dim a0d As New ArrayEx0
    Call a0d.Create
    If AssertHasError Then Else Exit Sub: Err.Clear
    Call a0d.Init(1).Create(2)
    If AssertHasNoError Then Else Exit Sub: Err.Clear
    Call a0d.Create(1)
    If AssertHasNoError Then Else Exit Sub: Err.Clear
    Call a0d.Create(array1d())
    If AssertHasError Then Else Exit Sub: Err.Clear
    Call a0d.Create(array2d())
    If AssertHasError Then Else Exit Sub: Err.Clear
       
    Dim a1d As New ArrayEx1
    Call a1d.Create
    If AssertHasError Then Else Exit Sub: Err.Clear
    Call a1d.Init(array1d()).Create(array1d())
    If AssertHasNoError Then Else Exit Sub: Err.Clear
    Call a1d.Create(1)
    If AssertHasError Then Else Exit Sub: Err.Clear
    Call a1d.Create(array1d())
    If AssertHasNoError Then Else Exit Sub: Err.Clear
    Call a1d.Create(array2d())
    If AssertHasError Then Else Exit Sub: Err.Clear
    
    Dim a2d As New ArrayEx2
    Call a2d.Create
    If AssertHasError Then Else Exit Sub: Err.Clear
    Call a2d.Init(array2d()).Create(array2d())
    If AssertHasNoError Then Else Exit Sub: Err.Clear
    Call a2d.Create(1)
    If AssertHasError Then Else Exit Sub: Err.Clear
    Call a2d.Create(array1d())
    If AssertHasError Then Else Exit Sub: Err.Clear
    Call a2d.Create(array2d())
    If AssertHasNoError Then Else Exit Sub: Err.Clear
End Sub


Sub GetValue_NoEception()
    On Error Resume Next

    Dim a2d As New ArrayEx2
    a2d.Init (array2d())

    If AssertEqual(False, IsNull(a2d.Value)) Then Else Exit Sub
    If AssertEqual(1 * 2, a2d.Value(1, 2)) Then Else Exit Sub
    If AssertEqual(5, UBound(a2d.Value(":", 1))) Then Else Exit Sub
    If AssertEqual(10, UBound(a2d.Value(2, ":"))) Then Else Exit Sub
    If AssertEqual(5, UBound(a2d.Value, 1)) Then Else Exit Sub
    If AssertEqual(10, UBound(a2d.Value, 2)) Then Else Exit Sub
    If AssertEqual(10, UBound(a2d.Value(rows:=2))) Then Else Exit Sub
    If AssertEqual(5, UBound(a2d.Value(cols:=1))) Then Else Exit Sub

    If AssertEqual("[1,2,3]", a2d.Extract(1, "1:3").ToString) Then Else Exit Sub
    If AssertEqual("[1,2,3]", a2d.Extract(1, "1 To 3").ToString) Then Else Exit Sub
    If AssertEqual("[1,2,3]", a2d.Extract(1, "1,2,3").ToString) Then Else Exit Sub
    If AssertEqual("[1,2,3;2,4,6]", a2d.Extract("1:2", "1:3").ToString) Then Else Exit Sub

    Call a2d.Extract(1, ":3").ToString
    If AssertHasError Then Else Exit Sub

    Dim a1d As New ArrayEx1
    a1d.Init (array1d())
    If AssertEqual(False, IsNull(a1d.Value)) Then Else Exit Sub
    If AssertEqual(1, a1d.Value(1)) Then Else Exit Sub
    If AssertEqual(2, a1d.Value("1:3")(2)) Then Else Exit Sub
    If AssertEqual(3, a1d.Value("1:3")(3)) Then Else Exit Sub
    If AssertEqual(3, UBound(a1d.Value("1:3"))) Then Else Exit Sub

    If AssertEqual("[1,2,3]", a1d.GetElements("1,2,3").ToString) Then Else Exit Sub
    If AssertEqual("[1,2,3]", a1d.GetElements("1 to 3").ToString) Then Else Exit Sub
    If AssertEqual("[1,2,3]", a1d.GetElements("1:3").ToString) Then Else Exit Sub
    If AssertEqual("[1,2,3,4,5]", a1d.GetElements(":").ToString) Then Else Exit Sub

    Dim a0d As New ArrayEx0
    a0d.Init (1)
    If AssertFalse(IsNull(a0d.Value)) Then Else Exit Sub
    If AssertEqual(1, a0d.Value) Then Else Exit Sub

End Sub

Sub SetValue_GetMethodIsCorrect_NoException()
    On Error Resume Next

    Dim a2d As New ArrayEx2
    a2d.Init (array2d())
 
    Dim a1d As New ArrayEx1
    a1d.Init (array1drow())
    Dim b1d As New ArrayEx1
    b1d.Init (array1d())
 
    Dim c0d As New ArrayEx0
    c0d.Init ("test")
 
    Call a2d.SetElement(1, 1, 10)
    If AssertEqual(10, a2d(1, 1)) Then Else Exit Sub
    Call a2d.SetRow(1, a1d)
    If AssertTrue(a2d.GetRow(1).Equals(a1d)) Then Else Exit Sub
    Call a2d.SetColumn(1, b1d)
    If AssertTrue(a2d.GetColumn(1).Equals(b1d)) Then Else Exit Sub

    Call a1d.SetElement(1, 10)
    If AssertEqual(10, a1d(1)) Then Else Exit Sub

    c0d = "sample"
    If AssertEqual("sample", c0d) Then Else Exit Sub
    
End Sub

Sub AddElement_NoException()
    On Error Resume Next
    Dim a2d As New ArrayEx2
    a2d.Init (array2d())

    If AssertEqual(6, a2d.AddRow(a2d.GetRow(1)).Ub(1)) Then Else Exit Sub
    If AssertEqual(11, a2d.AddColumn(a2d.GetColumn(1)).Ub(2)) Then Else Exit Sub

    Dim a1d As New ArrayEx1
    a1d.Init (array1d())
    Dim a0d As New ArrayEx0
    a0d.Init (10)
    If AssertEqual(6, a1d.AddElement(a0d).Ub) Then Else Exit Sub

End Sub

Sub Converts_NoException()
    On Error Resume Next

    Dim a0d As New ArrayEx0
    If AssertEqual("123", a0d.Init(123).ToString) Then Else Exit Sub

    Dim a2d As New ArrayEx2
    Dim collect As Collection
    Set collect = a2d.Init(array2d()).ToCollection
    If AssertEqual(5, collect.Count) Then Else Exit Sub
    If AssertEqual("[1,2,3]", collect.Item(1).Extract("1:3").ToString) Then Else Exit Sub

    Dim a1d As New ArrayEx1
    Set collect = a1d.Init(array1d()).ToCollection
    If AssertEqual(5, collect.Count) Then Else Exit Sub
    If AssertEqual(1, collect.Item(1).Value) Then Else Exit Sub

End Sub

Sub WrapFunctions_NoException()
    On Error Resume Next
    
    Dim a0d As New ArrayEx0
    If AssertEqual("fuga", a0d.Init("hoge").Replace("h", "f").Replace("oge", "uga")) Then Else Exit Sub
    If AssertEqual("323", a0d.Init(123).Replace("1", "3").ToString) Then Else Exit Sub
    If AssertEqual("45", a0d.Init("123456789").Left(5).Right(4).Mid(2).Mid(2, 2)) Then Else Exit Sub

End Sub


Sub DebugPrint_NoException()
    On Error Resume Next

    Dim a2d As New ArrayEx2
    a2d.Init (array2d())

    Call a2d.DebugPrint("{x}, {y}", "1,4", "no1 : {x}, no4:{y}").GetRow(1).DebugPrint("Test:{x}")
    Call a2d.Extract("1:3", "2:5").DebugPrintAll.GetColumn(1).DebugPrintAll.GetElement(3).DebugPrint
   
End Sub

Sub Array1Linqs_NoException()
    On Error Resume Next
    
    Dim a1d As New ArrayEx1
    a1d.Init (array1d())
    
    If AssertEqual(5, a1d.Count) Then Else Exit Sub
    If AssertTrue(a1d.Contains(4)) Then Else Exit Sub
    If AssertEqual(1, a1d.Min()) Then Else Exit Sub
    If AssertEqual(5, a1d.Max()) Then Else Exit Sub
        
    If AssertEqual(1, a1d.First()) Then Else Exit Sub
    If AssertEqual(5, a1d.Last()) Then Else Exit Sub
        
    If AssertEqual("[3,4,5]", a1d.Skip(2).ToString()) Then Else Exit Sub
    If AssertEqual("[1,2,3]", a1d.Take(3).ToString()) Then Else Exit Sub
    If AssertEqual("[1,2,3,4,5]", a1d.OrderBy().ToString()) Then Else Exit Sub
    If AssertEqual("[5,4,3,2,1]", a1d.OrderByDescending().ToString()) Then Else Exit Sub
    If AssertEqual("[5,4,3,2,1]", a1d.Reverse().ToString()) Then Else Exit Sub
    
    If AssertEqual("[4,5]", a1d.WhereEvaluated("x", "x > 3").ToString()) Then Else Exit Sub
    
    Call a1d.WhereEvaluated("x", "y>3")
    If AssertEqual(1003, Err.Number) Then Else Exit Sub
    Call a1d.WhereEvaluated("x", "x + 3")
    If AssertEqual(1004, Err.Number) Then Else Exit Sub
    
    If AssertEqual("[4,7,12,19,28]", a1d.SelectEvaluated("{x}", "{x}^2+ 3").ToString()) Then Else Exit Sub
    
    If AssertTrue(a1d.AllEvaluate("x", "x > 0 ")) Then Else Exit Sub
    If AssertFalse(a1d.AllEvaluate("x", "x > 1 ")) Then Else Exit Sub
    Call a1d.AllEvaluate("x", "x & Hello ")
    If AssertEqual(1003, Err.Number) Then Else Exit Sub
    Call a1d.AllEvaluate("x", "x + 0 ")
    If AssertEqual(1004, Err.Number) Then Else Exit Sub
    
    If AssertTrue(a1d.AnyEvaluate("x", "x > 4 ")) Then Else Exit Sub
    If AssertFalse(a1d.AnyEvaluate("x", "x > 5 ")) Then Else Exit Sub
    
    Dim a1dDistinct As New ArrayEx1
    a1dDistinct.Init (array1dHasDuplication)
    If AssertEqual("[1,2,4,5]", a1dDistinct.Distinct.ToString()) Then Else Exit Sub
    
End Sub

Sub Array2Linqs_NoException()
    On Error Resume Next
    
    Dim a2d As New ArrayEx2
    a2d.Init (array2d())
    
    Call a2d.WhereEvaluated("x,y,z", "1,2", "x+y=3")
    If AssertEqual(2002, Err.Number) Then Else Exit Sub
    Err.Clear
    Call a2d.WhereEvaluated("x,y", "1,2,3", "x+y=3")
    If AssertEqual(2002, Err.Number) Then Else Exit Sub
    Err.Clear
    
    If AssertEqual(1, a2d.WhereEvaluated("x,y", "1,2", "x+y=3").Count) Then Else Exit Sub
    If AssertEqual("[4,8,12,16,20]", a2d.SelectEvaluated("x,y", "1,3", "x+y").ToString) Then Else Exit Sub
    If AssertEqual("[1,2,3,4,5]", a2d.SelectEvaluated("x", "1", "x").ToString) Then Else Exit Sub
    
    If AssertTrue(a2d.AllEvaluate("x", "3", "x>=3")) Then Else Exit Sub
    If AssertTrue(a2d.AnyEvaluate("x", "3", "x>3")) Then Else Exit Sub
    
    If AssertEqual("[1,2,3,4,5,6,7,8,9,10]", a2d.First.ToString()) Then Else Exit Sub
    If AssertEqual("[5,10,15,20,25,30,35,40,45,50]", a2d.Last.ToString()) Then Else Exit Sub
    
    If AssertEqual(4, a2d.Skip(1).Count) Then Else Exit Sub
    If AssertEqual(3, a2d.Take(3).Count) Then Else Exit Sub
    If AssertEqual("[1,2,3,4,5]", a2d.OrderBy(1).GetColumn(1).ToString()) Then Else Exit Sub
    If AssertEqual("[5,4,3,2,1]", a2d.OrderByDescending(1).GetColumn(1).ToString()) Then Else Exit Sub
    If AssertEqual(a2d.ToString(), a2d.Reverse.Reverse.ToString()) Then Else Exit Sub
    
    Dim a2dDuplication As New ArrayEx2
    If AssertEqual("[1,2,4,5]", a2dDuplication.Init(array2dDuplication).Distinct(2).GetColumn(1).ToString) Then Else Exit Sub
    
End Sub

' ----------------------------------------------------------------------------------------------------------------------
Private Function RunTest(testName As String)
    Dim res As String
    Application.Run (testName)
    Debug.Print (IIf(Err.Number = 0, "OK", "NG") & ": " & testName & IIf(Err.Number = 0, "", ", " & Err.Number & ", " & Err.Source & ", " & Err.Description))
    Err.Clear
End Function

Private Function AssertTrue(condition) As Boolean
    AssertTrue = False
    On Error GoTo errCondition
    If condition = True Then
        AssertTrue = True
    Else
        Call Err.Raise(9001, "", "Should be True.")
    End If
    Exit Function
errCondition:
End Function

Private Function AssertFalse(condition) As Boolean
    AssertFalse = False
    On Error GoTo errCondition
    If condition = False Then
        AssertFalse = True
    Else
        Call Err.Raise(9002, "", "Should be False.")
    End If
    Exit Function
errCondition:
End Function

Private Function AssertEqual(expected, actual) As Boolean
    AssertEqual = False
    On Error GoTo errCondition
    If expected = actual Then
        AssertEqual = True
    Else
        Call Err.Raise(9003, "", "Should be equal. expected is " & expected & ", actual is " & actual)
    End If
    Exit Function
errCondition:
End Function

Private Function AssertNotEqual(expected, actual) As Boolean
    AssertNotEqual = False
    On Error GoTo errCondition
    If expected <> actual Then
        AssertNotEqual = True
    Else
        Call Err.Raise(9004, "", "Should not be equal. expected is " & expected & ", actual is " & actual)
    End If
    Exit Function
errCondition:
End Function

Private Function AssertHasError() As Boolean
    AssertHasError = IIf(Err.Number <> 0, True, False)
End Function

Private Function AssertHasNoError() As Boolean
    AssertHasNoError = IIf(Err.Number = 0, True, False)
End Function

