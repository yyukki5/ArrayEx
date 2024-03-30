Attribute VB_Name = "UnitTests"
Option Explicit

Private test_ As New UnitTest

Sub RunTests()
    Set test_ = New UnitTest

    test_.RegisterTest "Test_Ex012_Init"
    test_.RegisterTest "Test_Ex012_InitForce"
    test_.RegisterTest "Test_Ex012_Equals"
    test_.RegisterTest "Test_Ex012_GetValue"
    test_.RegisterTest "Test_Ex012_SetValue_GetMethod"
    test_.RegisterTest "Test_Ex012_AddElementRowColumn"
    test_.RegisterTest "Test_Ex012_ToString"
    test_.RegisterTest "Test_Ex012_ToCollection"
    test_.RegisterTest "Test_Ex012_ToRange"
    test_.RegisterTest "Test_Ex12_Extract"
    test_.RegisterTest "Test_Ex0_StringFunctions"
    test_.RegisterTest "Test_Ex1_Linqs"
    test_.RegisterTest "Test_Ex2_Linqs"

    test_.RegisterTest "Test_Core_Validate"
    test_.RegisterTest "Test_Core_RedimPreserve2"
    test_.RegisterTest "Test_Core_ShiftIndex"
    test_.RegisterTest "Test_Core_ConvertToDimensionN"
    test_.RegisterTest "Test_Core_IndexToArray"
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

Sub Test_Core_Validate()
    Dim ar As Variant, emptyVal As Variant, obj As Object
    
    On Error Resume Next
    
    ArrayExCore.Validate array1d
    test_.AssertHasNoError
    ArrayExCore.Validate Empty
    test_.AssertHasError
    ArrayExCore.Validate CVErr(xlErrNA)
    test_.AssertHasError
    ArrayExCore.Validate Null
    test_.AssertHasError
    ArrayExCore.Validate obj
    test_.AssertHasError
    ArrayExCore.Validate Nothing
    test_.AssertHasError
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


Public Sub Test_Core_IndexToArray()
    Dim reLong, rearr, reString1, reString2, reString3, reCollection
    Dim coll As New Collection
    Dim arr As New ArrayEx1
        
    rearr = ArrayExCore.IndexToArray([{1,2,3}], 0, 5)
    reLong = ArrayExCore.IndexToArray(1, 0, 5)
    reString1 = ArrayExCore.IndexToArray("1,2,3", 0, 5)
    reString2 = ArrayExCore.IndexToArray(":", 0, 5)
    reString3 = ArrayExCore.IndexToArray("1:3", 0, 5)
    
    coll.Add 1
    coll.Add 2
    coll.Add 3
    reCollection = ArrayExCore.IndexToArray(coll, 0, 5)
    With test_
        .AssertEqual arr.Init([{1,2,3}]), arr.Init(rearr)
        .AssertEqual arr.Init([{1}]), arr.Init(reLong)
        .AssertEqual arr.Init([{1,2,3}]), arr.Init(reString1)
        .AssertEqual arr.Init([{1,2,3,4,5}]), arr.Init(reString2)
        .AssertEqual arr.Init([{1,2,3}]), arr.Init(reString3)
        .AssertEqual arr.Init([{1,2,3}]), arr.Init(reCollection)
    
        On Error Resume Next
        Call ArrayExCore.IndexToArray(Null, 0, 5)
        .AssertHasError
        Call ArrayExCore.IndexToArray(Empty, 0, 5)
        .AssertHasError
        ArrayExCore.IndexToArray Nothing, 0, 5
        .AssertHasError
        Call ArrayExCore.IndexToArray([{1,2,3}], 0, 2)
        .AssertHasError
        Call ArrayExCore.IndexToArray([{1,2,3}], 2, 5)
        .AssertHasError
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
    
    txt = "1,2,3;4,5,6;a,b,c"
    str = ArrayExCore.TEXTSPLIT(txt, ",", ";")
    test_.AssertEqual 3, UBound(str, 1)
    test_.AssertEqual 3, UBound(str, 2)
    
    txt = "1,2:3::4;4,5,6/;;a,B,ab"
    str = ArrayExCore.TEXTSPLIT(txt, Array(",", ":", "b"), Array("/;", ";"), False)
    test_.AssertEqual 4, UBound(str, 1)
    test_.AssertEqual 5, UBound(str, 2)
    
    str = ArrayExCore.TEXTSPLIT(txt, Array(",", ":", "b"), Array("/;", ";"), True, 1, 123)
    test_.AssertEqual 3, UBound(str, 1)
    test_.AssertEqual 4, UBound(str, 2)
    test_.AssertEqual "a", str(3, 2)
    test_.AssertEqual 123, str(3, 4)
End Sub

Sub Test_Core_TAKE_DROP()
    Dim aex As New ArrayEx2
    Dim ar, arr
    ar = array2d
    
    arr = ArrayExCore.TAKE(ar, -3, 1)
    test_.AssertEqual "{3;4;5}", aex.Init(arr).ToString
    
    arr = ArrayExCore.TAKE(ar, 0, -2)
    test_.AssertTrue IsError(arr)
    arr = ArrayExCore.TAKE(ar, 1, 0)
    test_.AssertTrue IsError(arr)
    
    arr = ArrayExCore.DROP(ar, 1, 0)
    test_.AssertTrue (IsError(arr))
    arr = ArrayExCore.DROP(ar, 0, 1)
    test_.AssertTrue (IsError(arr))
        
End Sub

Sub Test_Ex012_Equals()
    Dim a0d1  As New ArrayEx0, a0d2 As New ArrayEx0
    Dim a1d1  As New ArrayEx1, a1d2 As New ArrayEx1
    Dim a2d1  As New ArrayEx2, a2d2 As New ArrayEx2
    
    a0d1.Init (1)
    a0d2.Init (1)
    a1d1.Init (array1d())
    a1d2.Init (array1d())
    a2d1.Init (array2d())
    a2d2.Init (array2d())
    
    test_.AssertEqual a0d1, a0d2
    test_.AssertEqual a1d1, a1d2
    test_.AssertEqual a2d1, a2d2
End Sub


Sub Test_Ex012_Init()
    Dim a0d As New ArrayEx0
    Dim a1d As New ArrayEx1
    Dim a2d As New ArrayEx2
    
    With test_
        On Error Resume Next
        a0d.Init (1)
        .AssertHasNoError
        a0d.Init (array1d)
        .AssertHasError
        a0d.Init (array2d)
        .AssertHasError
        a0d.Init (Empty)
        .AssertHasError
        a0d.Init (Null)
        .AssertHasError
        a0d.Init Nothing
        .AssertHasError
        
        a1d.Init (1)
        .AssertHasError
        a1d.Init (array1d())
        .AssertHasNoError
        a1d.Init (array2d())
        .AssertHasError
        a1d.Init (Empty)
        .AssertHasError
        a1d.Init (Null)
        .AssertHasError
        a1d.Init Nothing
        .AssertHasError
        
        a2d.Init (1)
        .AssertHasError
        a2d.Init (array1d())
        .AssertHasError
        a2d.Init (array2d())
        .AssertHasNoError
        a2d.Init (Empty)
        .AssertHasError
        a2d.Init (Null)
        .AssertHasError
        a2d.Init Nothing
        .AssertHasError
    End With

End Sub

Sub Test_Ex012_InitForce()
    Dim a0d As New ArrayEx0
    Dim a1d As New ArrayEx1
    Dim a2d As New ArrayEx2
    
    With test_
        .AssertEqual 1, a0d.InitForce(1).Value
        .AssertEqual 1, a0d.InitForce(array1d).Value
        .AssertEqual 1, a0d.InitForce(array2d).Value
        .AssertEqual 0, a1d.InitForce(1).Ub
        .AssertEqual 5, a1d.InitForce(array1d).Ub
        .AssertEqual 10, a1d.InitForce(array2d).Ub
        .AssertEqual 0, a2d.InitForce(1).Ub
        .AssertEqual 0, a2d.InitForce(array1d).Ub(1)
        .AssertEqual 10, a2d.InitForce(array2d).Ub(2)
    End With
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

Sub Test_Ex12_Extract()
    Dim a2d As New ArrayEx2
    Dim a1d As New ArrayEx1
    a2d.Init (array2d())
    a1d.Init (array1d())
    
    test_.AssertEqual "{1;2;3;4;5}", a2d.Extract(":", 1).ToString()
    test_.AssertEqual "{1,2,3;2,4,6}", a2d.Extract("1:2", "1:3").ToString()
    test_.AssertEqual "{1,2,3;2,4,6}", a2d.Extract("1 To 2", [{1,2,3}]).ToString()
    test_.AssertHasNoError
    
    On Error Resume Next
    a2d.Extract "a,b,c", "d,e,f"
    test_.AssertHasError
    a2d.Extract 100, -1
    test_.AssertHasError
    a2d.Extract 1, ":3"
    test_.AssertHasError
    On Error GoTo 0
    
    test_.AssertEqual "{1,2,3,4,5}", a1d.Extract(":").ToString
    test_.AssertEqual "{1,2,3}", a1d.Extract("1 : 3").ToString
    test_.AssertEqual "{1,2,3}", a1d.Extract("1 to 3").ToString
    test_.AssertEqual "{1,2,3}", a1d.Extract("1,2,3").ToString
    test_.AssertEqual "{1,2,3}", a1d.Extract([{1,2,3}]).ToString
    test_.AssertHasNoError
    
    On Error Resume Next
    a1d.Extract ("a,b,c")
    test_.AssertHasError
    a1d.Extract 10
    test_.AssertHasError
    On Error GoTo 0

End Sub

Sub Test_Ex012_GetValue()

    Dim a2d As New ArrayEx2
    a2d.Init (array2d())
    test_.AssertFalse IsNull(a2d.Value)
    test_.AssertEqual 1, a2d.Value(1, 1)

    Dim a1d As New ArrayEx1
    a1d.Init (array1d())
    test_.AssertFalse IsNull(a1d.Value)
    test_.AssertEqual 1, a1d.Value(1)
    test_.AssertEqual 2, a1d.Value(2)
    test_.AssertEqual 3, a1d.Value(3)
    test_.AssertEqual 5, UBound(a1d.Value)

    Dim a0d As New ArrayEx0
    a0d.Init (1)
    test_.AssertFalse IsNull(a0d.Value)
    test_.AssertEqual 1, a0d.Value

End Sub

Sub Test_Ex012_SetValue_GetMethod()
    Dim a2d As New ArrayEx2
    Dim a1d As New ArrayEx1
    Dim b1d As New ArrayEx1
    Dim c0d As New ArrayEx0
    
    a2d.Init (array2d())
    a1d.Init (array1drow())
    b1d.Init (array1d())
    c0d.Init ("test")
 
    Call a2d.SetElement(1, 1, 10)
    test_.AssertEqual 10, a2d.Value(1, 1)
    Call a2d.SetElement([{1,2,3}], "1 to 2", 2)
    test_.AssertEqual 2, a2d.Value(1, 1)
    test_.AssertEqual 2, a2d.Value(3, 2)
    
    Call a2d.SetRow(1, a1d)
    test_.AssertTrue a2d.GetRow(1).Equals(a1d)
    Call a2d.SetColumn(1, b1d)
    test_.AssertTrue a2d.GetColumn(1).Equals(b1d)

    Call a1d.SetElement(1, 10)
    test_.AssertEqual 10, a1d.Value(1)

    c0d.Value = "sample"
    test_.AssertEqual "sample", c0d.Value
End Sub

Sub Test_Ex012_AddElementRowColumn()
    Dim a2d As New ArrayEx2, a2dAdded As ArrayEx2
    Dim a1d As New ArrayEx1, a1dEmpty As ArrayEx1
    Dim a0d As New ArrayEx0
    
    a0d.Init (10)
    Set a1d = a1d.Init(array1d()).AddElement(a0d)
    a2d.Init (array2d())
    Set a2d = a2d.AddRow(a2d.GetRow(1)).AddColumn(a2d.GetColumn(1))
    
    test_.AssertEqual 6, a1d.Ub
    test_.AssertEqual 6, a2d.Ub(1)
    test_.AssertEqual 11, a2d.Ub(2)
End Sub

Sub Test_Ex012_ToString()
    Dim a0d As New ArrayEx0
    Dim a1d As New ArrayEx1
    Dim a2d As New ArrayEx2
    
    With test_
       .AssertEqual "abc", a0d.Init("abc").ToString()
       .AssertEqual "2023/10/01", a0d.Init(45200).ToString("yyyy/mm/dd")
       .AssertEqual "00123", a0d.Init(123).ToString("00000")
       .AssertEqual "", a0d.Init("").ToString()
        
       .AssertEqual "{1,2,3,4,5}", a1d.Init(array1d()).ToString()
       .AssertEqual "{1,2,3;2,4,6;3,6,9}", a2d.Init(array2d()).Extract("1:3", "1:3").ToString()
    End With
End Sub

Sub Test_Ex012_ToCollection()
    Dim collect1 As Collection, collect2 As Collection
    Dim a0d As New ArrayEx0, a1d As New ArrayEx1, a2d As New ArrayEx2
    
    Set collect1 = a1d.Init(array1d()).ToCollection()
    test_.AssertEqual 5, collect1.Count
    test_.AssertEqual 1, collect1.Item(1).Value
    
    Set collect2 = a2d.Init(array2d()).ToCollection(1)
    test_.AssertEqual 5, collect2.Count
    test_.AssertEqual "{1,2,3,4,5,6,7,8,9,10}", collect2.Item(1).ToString()
    
    Set collect2 = a2d.Init(array2d()).ToCollection(2)
    test_.AssertEqual 10, collect2.Count
    test_.AssertEqual "{2,4,6,8,10}", collect2(2).ToString()
    
    Set collect2 = a2d.Init(array2d()).ToCollection(0)
    test_.AssertEqual 50, collect2.Count
    Set collect2 = a2d.Init(array2d()).ToCollection(100)
    test_.AssertEqual 50, collect2.Count
End Sub

Sub Test_Ex012_ToRange()
    Dim a0d As New ArrayEx0, a1d As New ArrayEx1, a2d As New ArrayEx2
    Dim rng1 As Range, rng2 As Range
    Set rng1 = a1d.Init(array1d()).ToRange(Range("A1"))
    Set rng2 = a2d.Init(array2d()).ToRange(Range("A1"))
    test_.AssertEqual "$A$1:$E$1", rng1.Address
    test_.AssertEqual "$A$1:$J$5", rng2.Address
End Sub

Sub Test_Ex0_StringFunctions()
    Dim a0d As New ArrayEx0
    test_.AssertEqual "fuga", a0d.Init("hoge").Replace("h", "f").Replace("oge", "uga").ToString
    test_.AssertEqual "323", a0d.Init(123).Replace("1", "3").ToString
    test_.AssertEqual "45", a0d.Init("123456789").Left(5).Right(4).Mid(2).Mid(2, 2).Value
End Sub

Sub Test_Ex012_DebugPrint_NoTest()
    Dim a2d As New ArrayEx2
    
    a2d.Init (array2d())
    a2d.DebugPrint("{x}, {y}", "1,4", "no1 : {x}, no4:{y}").GetRow(1).DebugPrint ("Test:{x}")
    a2d.Extract("1:3", "2:5").DebugPrintAll.GetColumn(1).DebugPrintAll.GetElement(3).DebugPrint
End Sub

Sub Test_Ex1_Linqs()
    Dim a1d As New ArrayEx1
    a1d.Init (array1d())
        
    With test_
        .AssertTrue a1d.Contains(4)
        .AssertEqual 1, a1d.Min()
        .AssertEqual 5, a1d.Max()
        
        .AssertEqual 1, a1d.First().Value
        .AssertEqual 5, a1d.Last().Value
    
        .AssertEqual "{3,4,5}", a1d.Skip(2).ToString()
        .AssertEqual "{1,2,3}", a1d.TAKE(3).ToString()
        .AssertEqual "{1,2,3,4,5}", a1d.OrderByAscending().ToString()
        .AssertEqual "{5,4,3,2,1}", a1d.OrderByDescending().ToString()
        .AssertEqual "{5,4,3,2,1}", a1d.Reverse().ToString()
    
        .AssertEqual "{4,5}", a1d.WhereEvaluated("x", "x > 3").ToString()
        
        On Error Resume Next
        a1d.WhereEvaluated "x", "y>3"
        .AssertEqual 1003, Err.Number
        a1d.WhereEvaluated "x", "x + 3"
        .AssertEqual 1101, Err.Number
        a1d.WhereEvaluated "x", "x  & ""Hello"" "
        .AssertHasError
        On Error GoTo 0
        
        .AssertEqual "{4,7,12,19,28}", a1d.SelectEvaluated("{x}", "{x}^2+ 3").ToString()
        .AssertEqual "{1a,2a,3a,4a,5a}", a1d.SelectEvaluated("x", "x &""a""").ToString()
        .AssertEqual "{Abc,Abc,Abc,Abc,Abc}", a1d.SelectEvaluated("x", "proper(""abc"") ").ToString()
    
        .AssertTrue a1d.AllEvaluate("x", "x > 0 ")
        .AssertFalse a1d.AllEvaluate("x", "x > 1 ")
        
        On Error Resume Next
        a1d.AllEvaluate "x", "x & ""Hello"""
        .AssertEqual 1101, Err.Number
        a1d.AllEvaluate "x", "x + 0"
        .AssertEqual 1101, Err.Number
        On Error GoTo 0
        
        .AssertTrue a1d.AnyEvaluate("x", "x > 4 ")
        .AssertFalse a1d.AnyEvaluate("x", "x > 5 ")
    
        .AssertEqual "{1,2,4,5}", a1d.Init([{1,2,2,4,5}]).Distinct.ToString()
        .AssertEqual "{hoge,fuga}", a1d.Init([{"hoge","fuga","fuga","hoge" }]).Distinct.ToString()
    End With
    
End Sub

Sub Test_Ex2_Linqs()
    Dim a2d As New ArrayEx2
    a2d.Init (array2d())
        
    With test_
        On Error Resume Next
        a2d.WhereEvaluated "x,y,z", "1,2", "x+y=3"
        .AssertEqual 2002, Err.Number
        a2d.WhereEvaluated "x,y", "1,2,3", "x+y=3"
        .AssertEqual 2002, Err.Number
        On Error GoTo 0
        
        .AssertEqual 1, a2d.WhereEvaluated("x,y", "1,2", "x+y=3").Count
        .AssertEqual "{4,8,12,16,20}", a2d.SelectEvaluated("x,y", "1,3", "x+y").ToString
        .AssertEqual "{1,2,3,4,5}", a2d.SelectEvaluated("x", "1", "x").ToString
        .AssertFalse a2d.AllEvaluate("x", "2", "x>=3")
        .AssertTrue a2d.AnyEvaluate("x", "2", "x>3")
    
        .AssertEqual "{1,2,3,4,5,6,7,8,9,10}", a2d.First.ToString()
        .AssertEqual "{5,10,15,20,25,30,35,40,45,50}", a2d.Last.ToString()
    
        .AssertEqual 4, a2d.Skip(1).Count
        .AssertEqual 3, a2d.TAKE(3).Count
        .AssertEqual "{1,2,3,4,5}", a2d.OrderBy(1).GetColumn(1).ToString()
        .AssertEqual "{5,4,3,2,1}", a2d.OrderByDescending(1).GetColumn(1).ToString()
        .AssertEqual a2d.ToString(), a2d.Reverse.Reverse.ToString()
    
        Dim a2dDuplication As New ArrayEx2
        .AssertEqual "{1,2,4,5}", a2dDuplication.Init(array2dDuplication).Distinct(2).GetColumn(1).ToString
        
    End With
End Sub

' ----------------------------------------------------------------------------------------------------------------------

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
