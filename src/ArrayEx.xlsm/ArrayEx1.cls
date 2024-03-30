VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayEx1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ------------------------------------------------------------------------------------------------------------
' Init methods              : should be initialize by array.
' Get methods               : should return new instance.
' Set methods               : should return updated instance.
' Add methods               : should return new instance.
' Convert methods           : should return other type object.
' Linq methods              : should return new instance. some methods useing Evaluate()
' DebugPrint                : should return current instance. print current array value in immediate window.
' Wrap default functions    : should return new instance.
' ------------------------------------------------------------------------------------------------------------

' Fields -----------------------------------------------------------------------------------------------------
Private array_ As Variant

' Properties -------------------------------------------------------------------------------------------------
Public Property Get Value()
    Value = array_
End Property

Public Property Let Value(array1d)
    Init array1d
End Property

Public Property Get Ub()
    Ub = UBound(array_, 1)
End Property

Public Property Get Lb()
    Lb = LBound(array_, 1)
End Property

Public Property Get Count() As Long
    Count = UBound(array_, 1) - LBound(array_, 1) + 1
End Property

' Public Methods ---------------------------------------------------------------------------------------------
Public Function Init(array1d As Variant) As ArrayEx1
    ArrayExCore.Validate array1d
    If ArrayExCore.Rank(array1d) <> 1 Then Call Err.Raise(1101, "ArrayEx1.Init", "ArrayEx: Need 1d array")
    array_ = array1d
    Set Init = Me
End Function

Public Function InitForce(arr As Variant) As ArrayEx1
    array_ = ArrayExCore.ConvertToDimensionN(arr, 1)
    Set InitForce = Me
End Function

Public Function Equals(arrayEx1d As ArrayEx1) As Boolean
    Equals = False
    Dim array1d: array1d = arrayEx1d.Value
    If Lb <> LBound(array1d) Then Exit Function
    If Ub <> UBound(array1d) Then Exit Function
    Dim i As Long
    For i = Lb To Ub
        If array_(i) <> array1d(i) Then Exit Function
    Next i
    Equals = True
End Function

' Set
Public Function SetElement(index As Long, argVal) As ArrayEx1
    array_(index) = argVal
    Set SetElement = Me
End Function

Public Function Extract(indexes As Variant) As ArrayEx1
    Set Extract = ExtractCore(indexes)
End Function

Public Function GetElement(index As Long) As ArrayEx0
    Dim rearr As New ArrayEx0
    Set GetElement = rearr.Init(array_(index))
End Function

' Add
Function AddElement(element As ArrayEx0) As ArrayEx1
    Dim re: re = array_
    ReDim Preserve re(Lb To Ub + 1)
    re(Ub + 1) = element.Value
    
    Dim rearr As New ArrayEx1
    Set AddElement = rearr.Init(re)
End Function

' Convert
Public Function ToString() As String
    Dim s As String, i As Long, j As Long
    s = "{"
    For i = Lb To Ub
        s = s & CStr(array_(i)) & ","
    Next i
    s = Left(s, Len(s) - 1) & "}"
    ToString = s
End Function

Public Function ToRange(rng As Range) As Range
    rng.Resize(1, Count).Value = array_
    Set ToRange = rng.Resize(1, Count)
End Function

Public Function ToCollection() As Collection
    Dim coll As New Collection
    Dim arr As ArrayEx0
    Dim i As Long
    
    For i = Lb To Ub
        Set arr = New ArrayEx0
        coll.Add arr.Init(array_(i))
    Next i

    Set ToCollection = coll
End Function

' linqs
Public Function Contains(element As Variant) As Boolean
    Contains = True
    Dim i As Long
    For i = Lb To Ub
        If array_(i) = element Then Exit Function
    Next i
    Contains = False
End Function

Public Function First() As ArrayEx0
    Set First = GetElement(Lb)
End Function

Public Function Last() As ArrayEx0
    Set Last = GetElement(Ub)
End Function

Public Function Skip(skipIndex As Long) As ArrayEx1
    Set Skip = ExtractCore(Lb + skipIndex & ":" & Ub)
End Function

Public Function TAKE(takeIndex As Long) As ArrayEx1
    Set TAKE = ExtractCore(Lb & ":" & takeIndex + Lb - 1)
End Function

Public Function Distinct() As ArrayEx1
    Dim var As Variant
    Dim i As Long, s As String, sIndexes As String
    Dim coll As New Collection
    For i = Lb To Ub
        If IsDuplicated(coll, array_(i)) Then
        Else
            coll.Add array_(i)
            sIndexes = sIndexes & i & ","
        End If
    Next i
    If sIndexes = "" Then Exit Function
    sIndexes = Left(sIndexes, Len(sIndexes) - 1)
    Set Distinct = ExtractCore(sIndexes)
End Function

Public Function OrderByAscending() As ArrayEx1
    Dim re As New ArrayEx1
    Set OrderByAscending = re.Init(WorksheetFunction.Sort(array_))
End Function

Public Function OrderByDescending() As ArrayEx1
    Dim re As New ArrayEx1
    Set OrderByDescending = re.Init(WorksheetFunction.Sort(array_, , -1, True))
End Function

Public Function Reverse() As ArrayEx1
    Dim rearr: ReDim rearr(Lb To Ub)
    Dim i As Long
    For i = Lb To Ub
        rearr(i) = array_(Ub - i + 1)
    Next i
    Dim re As New ArrayEx1
    Set Reverse = re.Init(rearr)
End Function

Public Function AllEvaluate(argument As String, formula As String) As Boolean
    AllEvaluate = False
    Dim evaluation
    For Each evaluation In GetEvaluations(argument, formula)
        If VarType(evaluation) <> vbBoolean Then Err.Raise 1101, "ArrayEx1.AllEvaluate()", "Should evaluated value boolean, but not boolean: " & evaluation
        If evaluation = False Then Exit Function
    Next
    AllEvaluate = True
End Function

Public Function AnyEvaluate(argument As String, formula As String) As Boolean
    AnyEvaluate = True
    Dim evaluation
    For Each evaluation In GetEvaluations(argument, formula)
        If VarType(evaluation) <> vbBoolean Then Err.Raise 1101, "ArrayEx1.AllEvaluate()", "Should evaluated value boolean, but not boolean: " & evaluation
        If evaluation = True Then Exit Function
    Next
    AnyEvaluate = False
End Function

Public Function WhereEvaluated(argument As String, formula As String) As ArrayEx1
    Dim evaluations, indexes As String, i As Long
    evaluations = GetEvaluations(argument, formula)
    For i = Lb To Ub
        If VarType(evaluations(i)) <> vbBoolean Then Err.Raise 1101, "ArrayEx1.WhereEvaluated()", "Should evaluated value boolean, but not boolean: " & evaluations(i)
        If evaluations(i) Then indexes = indexes & i & ","
    Next i
    Set WhereEvaluated = ExtractCore(indexes)
End Function

Public Function SelectEvaluated(argument As String, formula As String) As ArrayEx1
    Dim evaluations
    evaluations = GetEvaluations(argument, formula)
    Dim rearr As New ArrayEx1
    Set SelectEvaluated = rearr.Init(evaluations)
End Function

' DebugPrint
Public Function DebugPrintAll() As ArrayEx1
    Debug.Print ToString
    Set DebugPrintAll = Me
End Function

Public Function DebugPrint(formulaWithArgumentX As String) As ArrayEx1
    Dim i&, s$
    For i = Lb To Ub
        s = Replace(formulaWithArgumentX, "{x}", array_(i))
        Debug.Print s
    Next i
    Set DebugPrint = Me
End Function

' Wrap default functions
Public Function Match(arg, Optional matchType As Long = 0) As Long
    Match = WorksheetFunction.Match(arg, array_, matchType)
End Function

Public Function Max() As Variant
    Max = WorksheetFunction.Max(array_)
End Function

Public Function Min() As Variant
    Min = WorksheetFunction.Min(array_)
End Function

Public Function Sum() As Variant
    Sum = WorksheetFunction.Sum(array_)
End Function

'Average
Public Function Average() As Variant
    Average = WorksheetFunction.Average(array_)
End Function


' Private Methods ------------------------------------------------------------------
Private Function ExtractCore(argIndexs As Variant) As ArrayEx1
    Dim re, rearr As New ArrayEx1
    re = ArrayExCore.ExtractCore(array_, argIndexs)
    Set ExtractCore = rearr.Init(re)
End Function

Private Function GetEvaluations(argument As String, formula As String) As Variant
    Dim i As Long, s As String, rearr, re
    ReDim rearr(Lb To Ub)
    For i = Lb To Ub
        s = Replace(formula, argument, array_(i))
        If ArrayExCore.TryEvaluate(s, re) Then rearr(i) = re _
            Else Err.Raise 1003, "ArrayEx1.GetEvaluations()", "String in evaluate() is error :" & s & "."
    Next i
    GetEvaluations = rearr
End Function

Private Function IsDuplicated(coll As Collection, val) As Boolean
    Dim col
    On Error GoTo errEqual
    IsDuplicated = True
    For Each col In coll
        If col = val Then Exit Function:
    Next
    IsDuplicated = False
errEqual:
End Function

