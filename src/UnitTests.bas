Attribute VB_Name = "UnitTests"
Option Explicit

Private test_ As New UnitTest

Sub RunTests()

    Set test_ = New UnitTest

    test_.RegisterTest "Init_NoException"
    test_.RegisterTest "Create_NoException"
    test_.RegisterTest "GetValue_NoEception"
    test_.RegisterTest "SetValue_GetMethodIsCorrect_NoException"
    test_.RegisterTest "AddElement_NoException"
    test_.RegisterTest "Converts_NoException"
    '    test_.RegisterTest "DebugPrint_NoException"
    test_.RegisterTest "WrapFunctions_NoException"
    test_.RegisterTest "Array1Linqs_NoException"
    test_.RegisterTest "Array2Linqs_NoException"
    

'    test_.RegisterTest "Test_Test"
    test_.RegisterTest "Test_Core_RedimPreserve2"
    test_.RegisterTest "Test_Core_ShiftIndex"
    test_.RegisterTest "Test_Core_ConvertToDimensionN"
    test_.RegisterTest "Test_Core_TEXTAFTER"
    test_.RegisterTest "Test_Core_TEXTBEFORE"
    test_.RegisterTest "Test_Core_TEXTSPLIT"
    test_.RegisterTest "Test_Core_TOCOL"
    test_.RegisterTest "Test_Core_VHSTACK"
    test_.RegisterTest "Test_Core_WRAPCOLS"
    test_.RegisterTest "Test_Core_WRAPROW"
    test_.RegisterTest "Test_Core_EXPAND"
    
    test_.RunTests
End Sub



Sub Test_Test()
    Set test_ = New UnitTest
    test_.AssertTrue True
    test_.AssertTrue False
    test_.AssertFalse False
    test_.AssertFalse True
    test_.AssertEqual 1, 1
    test_.AssertEqual 1, 2
    test_.AssertNotEqual 1, 2
    test_.AssertNotEqual 1, 1
    
    On Error Resume Next
    Err.Raise 9001
    test_.AssertHasError
    Err.Raise 9001
    test_.AssertHasNoError
    
    Err.Clear
    test_.AssertHasError
    Err.Clear
    test_.AssertHasNoError
    
End Sub

Sub Test_Core_GetSpllingRange()
    ' no test
End Sub

Sub Test_Core_RedimPreserve2()
    Dim arr, rearr
    arr = array2d
    
    rearr = ArrayExCore.RedimPreserve2(arr, 1, 3, 2, 5)
    
    test_.AssertEqual 1, LBound(rearr, 1)
    test_.AssertEqual 3, UBound(rearr, 1)
    test_.AssertEqual 2, LBound(rearr, 2)
    test_.AssertEqual 5, UBound(rearr, 2)
    test_.AssertEqual arr(1, 1), rearr(1, 2)
End Sub

Sub Test_Core_ShiftIndex()
    Dim arr, rearr2, rearr1, rearrerr1
    
    rearr1 = ArrayExCore.ShiftIndex((array1d), 1)
    rearr2 = ArrayExCore.ShiftIndex((array2d), -1, 1)
    
    test_.AssertEqual rearr1(2), array1d(1)
    test_.AssertEqual rearr1(3), array1d(2)
    test_.AssertEqual rearr1(4), array1d(3)
    test_.AssertEqual rearr1(5), array1d(4)
    test_.AssertEqual rearr1(6), array1d(5)
    test_.AssertEqual rearr2(1, 2), 2
    
    On Error Resume Next
    rearrerr1 = ArrayExCore.ShiftIndex((array1d), -3)
    test_.AssertHasError
    rearrerr1 = ArrayExCore.ShiftIndex((array2d), 0, -3)
    test_.AssertHasError
    
End Sub


Public Sub Test_Core_ConvertToDimensionN()
    Dim rea1, rea0, rea2
    Dim a2, a1, a0
    a0 = 1
    a1 = array1d
    a2 = array2d
    
    With ArrayExCore
        rea2 = .ConvertToDimensionN(a2, 2)
        rea1 = .ConvertToDimensionN(a2, 1)
        rea0 = .ConvertToDimensionN(a2, 0)
        
        test_.AssertEqual LBound(a2, 1), LBound(rea2, 1)
        test_.AssertEqual LBound(a2, 2), LBound(rea2, 2)
        test_.AssertEqual UBound(a2, 1), UBound(rea2, 1)
        test_.AssertEqual UBound(a2, 2), UBound(rea2, 2)
        test_.AssertEqual LBound(a2, 2), LBound(rea1, 1)
        test_.AssertEqual UBound(a2, 2), UBound(rea1, 1)
        test_.AssertEqual rea0, a2(LBound(a2, 1), LBound(a2, 1))
    
        rea2 = .ConvertToDimensionN(a1, 2)
        rea1 = .ConvertToDimensionN(a1, 1)
        rea0 = .ConvertToDimensionN(a1, 0)
        
        test_.AssertEqual 0, LBound(rea2, 1)
        test_.AssertEqual 0, UBound(rea2, 1)
        test_.AssertEqual LBound(a1, 1), LBound(rea2, 2)
        test_.AssertEqual UBound(a1, 1), UBound(rea2, 2)
        
        rea2 = .ConvertToDimensionN(a0, 2)
        rea1 = .ConvertToDimensionN(a0, 1)
        rea0 = .ConvertToDimensionN(a0, 0)
    End With
End Sub

Sub Test_Core_TEXTAFTER()
    Dim txt As String, re
    txt = "abc,dEf,ghi"
    
    re = ArrayExCore.TEXTAFTER(txt, ",")
    test_.AssertEqual "dEf,ghi", re
    re = ArrayExCore.TEXTAFTER(txt, ",", 2)
    test_.AssertEqual "ghi", re
    re = ArrayExCore.TEXTAFTER(txt, ",", -1)
    test_.AssertEqual "ghi", re
    re = ArrayExCore.TEXTAFTER(txt, "E", 1, 1)
    test_.AssertEqual "f,ghi", re
    re = ArrayExCore.TEXTAFTER(txt, "e", 1, 0)
    test_.AssertEqual CVErr(xlErrNA), re
    re = ArrayExCore.TEXTAFTER(txt, "e", 1, 0, , 123)
    test_.AssertEqual 123, re
End Sub
    
Sub Test_Core_TEXTBEFORE()
    Dim txt As String, re
    txt = "abc,dEf,ghi"
    
    re = ArrayExCore.TEXTBEFORE(txt, ",")
    test_.AssertEqual "abc", re
    re = ArrayExCore.TEXTBEFORE(txt, ",", 2)
    test_.AssertEqual "dEf", re
    re = ArrayExCore.TEXTBEFORE(txt, ",", -1)
    test_.AssertEqual "ghi", re
    re = ArrayExCore.TEXTBEFORE(txt, "E", 1, 1)
    test_.AssertEqual "abc,d", re
    re = ArrayExCore.TEXTBEFORE(txt, "e", 1, 0)
    test_.AssertEqual CVErr(xlErrNA), re
    re = ArrayExCore.TEXTBEFORE(txt, "e", 1, 0, , 123)
    test_.AssertEqual 123, re
End Sub

Sub Test_Core_TEXTSPLIT()
    Dim str, txt As String
    
    txt = textWithColumnRowDelimiters
    str = ArrayExCore.TEXTSPLIT(txt, ",", ";")
    test_.AssertEqual 3, UBound(str, 1)
    test_.AssertEqual 3, UBound(str, 2)
    
    txt = textWithColumnRowDelimiters2
    str = ArrayExCore.TEXTSPLIT(txt, Array(",", ":", "b"), Array("/;", ";"), False)
    test_.AssertEqual 4, UBound(str, 1)
    test_.AssertEqual 5, UBound(str, 2)
    
    str = ArrayExCore.TEXTSPLIT(txt, Array(",", ":", "b"), Array("/;", ";"), True, 1, 123)
    test_.AssertEqual 3, UBound(str, 1)
    test_.AssertEqual 4, UBound(str, 2)
    test_.AssertEqual "a", str(3, 2)
    test_.AssertEqual 123, str(3, 4)
End Sub



Sub abcd()
On Error Resume Next
    Dim a1 As New ArrayEx1
    Dim arr
    arr = Array(1, 2, 3, 4, 5)
    a1.Init (arr)
    test_.AssertHasNoError
    
    Set arr = Range("A1:A4")
    a1.Init arr
    test_.AssertHasError
    
    ReDim arr(1 To 5)
    arr(1) = 1
    arr(3) = 3
    arr(5) = 5
    a1.Init arr
    With test_
        .AssertEqual 0, a1.lb
        .AssertEqual 4, a1.ub
    End With
    
    a1.Init (3)
    With test_
        .AssertEqual 0, a1.lb
        .AssertTrue VarType(a1.Value) > vbArray
    End With
    
    
End Sub

Sub Test_Core_TAKE_DROP()
    Dim aex As New ArrayEx2
    Dim ar, arr
    ar = array2d
    
    arr = ArrayExCore.TAKE(ar, -3, 1)
    test_.AssertEqual "[3;4;5]", aex.Init(arr).ToString
    
    arr = ArrayExCore.TAKE(ar, 0, -2)
    test_.AssertTrue IsError(arr)
    arr = ArrayExCore.TAKE(ar, 1, 0)
    test_.AssertTrue IsError(arr)
    
    
    arr = ArrayExCore.DROP(ar, 1, 0)
    test_.AssertTrue (IsError(arr))
    arr = ArrayExCore.DROP(ar, 0, 1)
    test_.AssertTrue (IsError(arr))
    
    
    Call ArrayExCore.DROP(ar, 2)
    arr = ArrayExCore.DROP(ar, , 1)
    Call ArrayExCore.DROP(ar, -1, -1)
    

End Sub

Sub Init_NoError()
    Dim arr
    Dim a2 As New ArrayEx2
    arr = array2d()
    a2.Init arr
    
    With test_
        .AssertEqual 1, a2.lb(1)
        .AssertEqual 1, a2.lb(2)
    End With

    a2.Init array1d()
    test_.AssertEqual 0, a2.lb
    
    a2.Init 3
    test_.AssertEqual 0, a2.lb
    
    
    Dim re
    a2.Init array2d()
    re = a2.GetRow(2).Value
    test_.AssertEqual 10, UBound(re)
End Sub



Sub Test_Core_EXPAND()
    Dim a2, a2_, a1, a0
    a2 = array2d_1
    
    Dim arr, arr2
    arr = ArrayExCore.EXPAND(a2, 15, 15)
    arr2 = ArrayExCore.EXPAND(a2, 15, 15, 123)
    
    test_.AssertEqual 15, UBound(arr, 1)
    test_.AssertEqual 14, UBound(arr, 2)
    test_.AssertEqual CVErr(xlErrNA), arr(15, 14)
    test_.AssertEqual 123, arr2(15, 14)
End Sub

Sub Test_Core_VHSTACK()
    Dim a2, a2_, a1, a0, arr
    a2 = array2d
    a2_ = array2d_1
    a1 = array1d
    a0 = 1
    
    arr = ArrayExCore.VSTACK(a2, a2_, a1, a0)
    test_.AssertEqual 10, UBound(arr, 1)
    test_.AssertEqual 13, UBound(arr, 2)
    test_.AssertEqual 1, LBound(arr, 1)
    test_.AssertEqual 1, LBound(arr, 2)
    test_.AssertEqual 1, arr(6, 1)
    test_.AssertEqual 1, arr(9, 1)
    test_.AssertEqual 1, arr(10, 1)
    test_.AssertEqual CVErr(xlErrNA), arr(10, 13)
    
    arr = ArrayExCore.HSTACK(a2, a2_, a1, a0)
    test_.AssertEqual 5, UBound(arr, 1)
    test_.AssertEqual 29, UBound(arr, 2)
    test_.AssertEqual 1, LBound(arr, 1)
    test_.AssertEqual 1, LBound(arr, 2)
    test_.AssertEqual 1, arr(1, 11)
    test_.AssertEqual 1, arr(1, 24)
    test_.AssertEqual 1, arr(1, 29)
    test_.AssertEqual CVErr(xlErrNA), arr(2, 24)
    
End Sub

Sub Test_Core_TOCOL()
    Dim a2, arr
    a2 = array2d_2x3
    
    arr = ArrayExCore.TOCOL(a2)
    test_.AssertEqual 6, UBound(arr, 1)
    test_.AssertEqual 1, LBound(arr, 2)
    test_.AssertEqual 2, arr(2, 1)
    
    arr = ArrayExCore.TOCOL(a2, False)
    test_.AssertEqual 2, arr(2, 1)
    arr = ArrayExCore.TOCOL(a2, , True)
    test_.AssertEqual 4, arr(2, 1)
End Sub


Sub Test_Core_WRAPCOLS()
    Dim a1, arr
    a1 = array1d
     
    arr = ArrayExCore.WRAPCOLS(a1, 2)
    test_.AssertEqual CVErr(xlErrNA), arr(2, 3)
    arr = ArrayExCore.WRAPCOLS(a1, 2, 123)
    test_.AssertEqual 123, arr(2, 3)
    arr = ArrayExCore.WRAPCOLS(a1, 0)
    test_.AssertEqual CVErr(xlErrNum), arr
End Sub

Sub Test_Core_WRAPROW()
    Dim a1, arr
    a1 = array1d
     
    arr = ArrayExCore.WRAPROW(a1, 2)
    test_.AssertEqual CVErr(xlErrNA), arr(3, 2)
    arr = ArrayExCore.WRAPROW(a1, 2, 123)
    test_.AssertEqual 123, arr(3, 2)
    arr = ArrayExCore.WRAPROW(a1, 0)
    test_.AssertEqual CVErr(xlErrNum), arr
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

Private Function array2d_1() As Variant
    Dim a
    ReDim a(1 To 3, 0 To 12)
    Dim i, j
    For i = 1 To 3
        For j = 0 To 12
            a(i, j) = i + j
        Next j
    Next i
    array2d_1 = a
End Function

Private Function array2d_2x3() As Variant
    Dim a, cnt As Long
    ReDim a(1 To 2, 1 To 3)
    Dim i, j
    For i = 1 To 2
        For j = 1 To 3
            cnt = cnt + 1
            a(i, j) = cnt
        Next j
    Next i
    array2d_2x3 = a
End Function

Private Function array2dDuplication() As Variant
    Dim a: a = array2d
    a(3, 2) = 4
    array2dDuplication = a
End Function

Private Function textWithColumnRowDelimiters() As String
    textWithColumnRowDelimiters = "1,2,3;4,5,6;a,b,c"
End Function
Private Function textWithColumnRowDelimiters2() As String
    textWithColumnRowDelimiters2 = "1,2:3::4;4,5,6/;;a,B,ab"
End Function

Sub Test_2D_ShiftIndex_()
    Dim a2d As New ArrayEx2, a2d1 As ArrayEx2, a2d2 As ArrayEx2
    a2d.Init array2d()
    
    Set a2d1 = a2d.ShiftIndex(-1, -1)
    Set a2d2 = a2d.ShiftIndex(1, 2)
    
    With test_
        .AssertEqual 0, a2d1.lb
        .AssertEqual 0, a2d1.lb(2)
        .AssertEqual 2, a2d2.lb(1)
        .AssertEqual 3, a2d2.lb(2)
    End With
End Sub

Sub Init_NoException()
    On Error Resume Next

    With test_
        Dim a0d As New ArrayEx0
        a0d.Init (1)
        .AssertHasNoError
        a0d.Init (array1d)
        .AssertHasError
        a0d.Init (array2d)
        .AssertHasError
        
        Dim a1d As New ArrayEx1
        a1d.Init (1)
        .AssertHasError
        a1d.Init (array1d)
        .AssertHasNoError
        a1d.Init (array2d)
        .AssertHasError
        
        Dim a2d As New ArrayEx2
        a2d.Init (1)
        .AssertHasError
        a2d.Init (array1d())
        .AssertHasError
        .AssertEqual 0, a2d.lb(1)
        a2d.Init (array2d())
        .AssertHasNoError
        
        .AssertTrue (a2d.Equals(array2d()))
    End With

End Sub

Sub Create_NoException()
    On Error Resume Next
    
    Dim a0d As New ArrayEx0
    a0d.Create
    test_.AssertHasError
    a0d.Init(1).Create (2)
    test_.AssertHasNoError
    a0d.Create (1)
    test_.AssertHasNoError
    a0d.Create (array1d())
    test_.AssertHasError
    a0d.Create (array2d())
    test_.AssertHasError
       
    Dim a1d As New ArrayEx1
    a1d.Create
    test_.AssertHasError
    a1d.Init(array1d()).Create (array1d())
    test_.AssertHasNoError
    a1d.Create (1)
    test_.AssertHasError
    a1d.Create (array1d())
    test_.AssertHasNoError
    a1d.Create (array2d())
    test_.AssertHasError
    
    Dim a2d As New ArrayEx2
    a2d.Create
    test_.AssertHasError
    a2d.Init(array2d()).Create (array2d())
    test_.AssertHasNoError
    a2d.Create (1)
    test_.AssertHasError
    a2d.Create (array1d())
    test_.AssertHasError
    a2d.Create (array2d())
    test_.AssertHasNoError
End Sub


Sub GetValue_NoEception()
    On Error Resume Next

    Dim a2d As New ArrayEx2
    a2d.Init (array2d())
    test_.AssertFalse IsNull(a2d.Value)
    test_.AssertEqual "[1,2,3]", a2d.Extract(1, "1:3").ToString()
    test_.AssertEqual "[1,2,3]", a2d.Extract(1, "1 To 3").ToString()
    test_.AssertEqual "[1,2,3]", a2d.Extract(1, "1,2,3").ToString()
    test_.AssertEqual "[1,2,3;2,4,6]", a2d.Extract("1:2", "1:3").ToString()

    a2d.Extract(1, ":3").ToString
    test_.AssertHasError

    Dim a1d As New ArrayEx1
    a1d.Init (array1d())
    test_.AssertFalse IsNull(a1d.Value)
    test_.AssertEqual 1, a1d.Value(1)
    test_.AssertEqual 2, a1d.Value("1:3")(2)
    test_.AssertEqual 3, a1d.Value("1:3")(3)
    test_.AssertEqual 3, UBound(a1d.Value("1:3"))
    test_.AssertEqual "[1,2,3]", a1d.GetElements("1,2,3").ToString
    test_.AssertEqual "[1,2,3]", a1d.GetElements("1 to 3").ToString
    test_.AssertEqual "[1,2,3]", a1d.GetElements("1:3").ToString
    test_.AssertEqual "[1,2,3,4,5]", a1d.GetElements(":").ToString

    Dim a0d As New ArrayEx0
    a0d.Init (1)
    test_.AssertFalse IsNull(a0d.Value)
    test_.AssertEqual 1, a0d.Value

End Sub

Sub SetValue_GetMethodIsCorrect_NoException()
    On Error Resume Next

    Dim a2d As New ArrayEx2
    Dim a1d As New ArrayEx1
    Dim b1d As New ArrayEx1
    Dim c0d As New ArrayEx0
    
    a2d.Init (array2d())
    a1d.Init (array1drow())
    b1d.Init (array1d())
    c0d.Init ("test")
 
    Call a2d.SetElement(1, 1, 10)
    test_.AssertEqual 10, a2d(1, 1)
    Call a2d.SetRow(1, a1d)
    test_.AssertTrue a2d.GetRow(1).Equal(a1d)
    Call a2d.SetColumn(1, b1d)
    test_.AssertTrue a2d.GetColumn(1).Equal(b1d)

    Call a1d.SetElement(1, 10)
    test_.AssertEqual 10, a1d(1)

    c0d.Value = "sample"
    test_.AssertEqual "sample", c0d.Value
End Sub

Sub AddElement_NoException()
    On Error Resume Next
    Dim a2d As New ArrayEx2
    Dim a1d As New ArrayEx1
    Dim a0d As New ArrayEx0
    
    a2d.Init (array2d())
    a1d.Init (array1d())
    a0d.Init (10)
    
    test_.AssertEqual 6, a2d.AddRow(a2d.GetRow(1)).ub(1)
    test_.AssertEqual 11, a2d.AddColumn(a2d.GetColumn(1)).ub(2)
    test_.AssertEqual 6, a1d.AddElement(a0d).ub
End Sub

Sub Converts_NoException()
    Dim collect1 As Collection, collect2 As Collection
    Dim a0d As New ArrayEx0
    Dim a1d As New ArrayEx1
    Dim a2d As New ArrayEx2
    
    Set collect1 = a1d.Init(array1d()).ToCollection
    Set collect2 = a2d.Init(array2d()).ToCollection
    
    test_.AssertEqual "123", a0d.Init(123).ToString
    test_.AssertEqual 5, collect2.Count
    test_.AssertEqual "[1,2,3]", collect2.Item(1).Extract("1:3").ToString
    test_.AssertEqual 5, collect1.Count
    test_.AssertEqual 1, collect1.Item(1).Value
End Sub

Sub Test_Array0_StringsFunctions_NoException()
    Dim a0d As New ArrayEx0
    test_.AssertEqual "fuga", a0d.Init("hoge").Replace("h", "f").Replace("oge", "uga").ToString
    test_.AssertEqual "323", a0d.Init(123).Replace("1", "3").ToString
    test_.AssertEqual "45", a0d.Init("123456789").Left(5).Right(4).Mid(2).Mid(2, 2).Value
End Sub

Sub Test_012_DebugPrint_NoTest()
    Dim a2d As New ArrayEx2
    Dim a1d As New ArrayEx1
    Dim a0d As New ArrayEx0
    
    a2d.Init (array2d())
    a2d.DebugPrint("{x}, {y}", "1,4", "no1 : {x}, no4:{y}").GetRow(1).DebugPrint ("Test:{x}")
    a2d.Extract("1:3", "2:5").DebugPrintAll.GetColumn(1).DebugPrintAll.GetElement(3).DebugPrint
End Sub

Sub Array1Linqs_NoException()
    On Error Resume Next
    
    Dim a1d As New ArrayEx1
    a1d.Init (array1d())
    
With test_
    .AssertEqual 5, a1d.Count
    .AssertTrue a1d.Contains(4)
    .AssertEqual 1, a1d.Min()
    .AssertEqual 5, a1d.Max()
    
    .AssertEqual 1, a1d.First().Value
    .AssertEqual 5, a1d.Last().Value

    .AssertEqual "[3,4,5]", a1d.Skip(2).ToString()
    .AssertEqual "[1,2,3]", a1d.TAKE(3).ToString()
    .AssertEqual "[1,2,3,4,5]", a1d.OrderBy().ToString()
    .AssertEqual "[5,4,3,2,1]", a1d.OrderByDescending().ToString()
    .AssertEqual "[5,4,3,2,1]", a1d.Reverse().ToString()

    .AssertEqual "[4,5]", a1d.WhereEvaluated("x", "x > 3").ToString()
    a1d.WhereEvaluated "x", "y>3"
    .AssertEqual 1003, Err.Number
    a1d.WhereEvaluated "x", "x + 3"
    .AssertHasNoError
    
    .AssertEqual "[4,7,12,19,28]", a1d.SelectEvaluated("{x}", "{x}^2+ 3").ToString()

    .AssertTrue a1d.AllEvaluate("x", "x > 0 ")
    .AssertFalse a1d.AllEvaluate("x", "x > 1 ")
    a1d.AllEvaluate "x", "x & Hello "
    .AssertEqual 1003, Err.Number
    a1d.AllEvaluate "x", "x + 0"
        
    .AssertTrue a1d.AnyEvaluate("x", "x > 4 ")
    .AssertFalse a1d.AnyEvaluate("x", "x > 5 ")

    Dim a1dDistinct As New ArrayEx1
    a1dDistinct.Init (array1dHasDuplication)
    .AssertEqual "[1,2,4,5]", a1dDistinct.Distinct.ToString()
End With

    
End Sub

Sub Array2Linqs_NoException()
    On Error Resume Next
    
    Dim a2d As New ArrayEx2
    a2d.Init (array2d())
        
    With test_
        a2d.WhereEvaluated "x,y,z", "1,2", "x+y=3"
        .AssertEqual 2002, Err.Number
        a2d.WhereEvaluated "x,y", "1,2,3", "x+y=3"
        .AssertEqual 2002, Err.Number
        
        .AssertEqual 1, a2d.WhereEvaluated("x,y", "0,1", "x+y=3").Count
        .AssertEqual "[4,8,12,16,20]", a2d.SelectEvaluated("x,y", "0,2", "x+y").ToString
        .AssertEqual "[1,2,3,4,5]", a2d.SelectEvaluated("x", "0", "x").ToString
        .AssertTrue a2d.AllEvaluate("x", "2", "x>=3")
        .AssertTrue a2d.AnyEvaluate("x", "2", "x>3")
    
        .AssertEqual "[1,2,3,4,5,6,7,8,9,10]", a2d.First.ToString()
        .AssertEqual "[5,10,15,20,25,30,35,40,45,50]", a2d.Last.ToString()
    
        .AssertEqual 4, a2d.Skip(1).Count
        .AssertEqual 3, a2d.TAKE(3).Count
        .AssertEqual "[1,2,3,4,5]", a2d.OrderBy(1).GetColumn(0).ToString()
        .AssertEqual "[5,4,3,2,1]", a2d.OrderByDescending(1).GetColumn(0).ToString()
        .AssertEqual a2d.ToString(), a2d.Reverse.Reverse.ToString()
    
        Dim a2dDuplication As New ArrayEx2
        .AssertEqual "[1,2,4,5]", a2dDuplication.Init(array2dDuplication).Distinct(1).GetColumn(0).ToString
    End With
End Sub

' ----------------------------------------------------------------------------------------------------------------------
'Private Function RunTest(testName As String)
'    Dim res As String
'    Application.Run (testName)
'    Debug.Print (IIf(Err.Number = 0, "OK", "NG") & ": " & testName & IIf(Err.Number = 0, "", ", " & Err.Number & ", " & Err.Source & ", " & Err.Description))
'    Err.Clear
'End Function
'
'Private Function AssertTrue(condition) As Boolean
'    AssertTrue = False
'    On Error GoTo errCondition
'    If condition = True Then
'        AssertTrue = True
'    Else
'        Call Err.Raise(9001, "", "Should be True.")
'    End If
'    Exit Function
'errCondition:
'End Function
'
'Private Function AssertFalse(condition) As Boolean
'    AssertFalse = False
'    On Error GoTo errCondition
'    If condition = False Then
'        AssertFalse = True
'    Else
'        Call Err.Raise(9002, "", "Should be False.")
'    End If
'    Exit Function
'errCondition:
'End Function
'
'Private Function AssertEqual(expected, actual) As Boolean
'    AssertEqual = False
'    On Error GoTo errCondition
'    If expected = actual Then
'        AssertEqual = True
'    Else
'        Call Err.Raise(9003, "", "Should be equal. expected is " & expected & ", actual is " & actual)
'    End If
'    Exit Function
'errCondition:
'End Function
'
'Private Function AssertNotEqual(expected, actual) As Boolean
'    AssertNotEqual = False
'    On Error GoTo errCondition
'    If expected <> actual Then
'        AssertNotEqual = True
'    Else
'        Call Err.Raise(9004, "", "Should not be equal. expected is " & expected & ", actual is " & actual)
'    End If
'    Exit Function
'errCondition:
'End Function
'
'Private Function AssertHasError() As Boolean
'    AssertHasError = IIf(Err.Number <> 0, True, False)
'End Function
'
'Private Function AssertHasNoError() As Boolean
'    AssertHasNoError = IIf(Err.Number = 0, True, False)
'End Function

