VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayEx2"
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

Public Property Get ub(Optional direction As Long = 1) As Long
    ub = UBound(array_, direction)
End Property

Public Property Get lb(Optional direction As Long = 1) As Long
    lb = LBound(array_, direction)
End Property

Public Property Get Count(Optional direction As Long = 1) As Long
    Count = UBound(array_, direction) - LBound(array_, direction) + 1
End Property

Public Property Get Rank() As Long
    Rank = ArrayExCore.Rank(array_)
End Property

' Public Methods ---------------------------------------------------------------------------------------------
Public Function Init(array2d As Variant) As ArrayEx2
    ArrayExCore.Validate array2d
    If ArrayExCore.Rank(array2d) <> 2 Then Call Err.Raise(2101, "ArrayEx2 Init", "ArrayEx: Need 2d array")
    array_ = array2d
    Set Init = Me
End Function

Public Function InitForce(arr As Variant) As ArrayEx2
    array_ = ArrayExCore.ConvertToDimensionN(arr, 2)
    Set InitForce = Me
End Function

'todo delete ???
Public Function Create(Optional array2d As Variant) As ArrayEx2
    If IsMissing(array2d) = True And IsEmpty(array_) Then Call Err.Raise(2105, "ArrayEx2 Create", "ArrayEx: Not initialized")
    Dim rearr As New ArrayEx2
    If IsMissing(array2d) Then
        Set Create = rearr.Init(array_)
    Else
        Set Create = rearr.Init(array2d)
    End If
End Function

Public Function Equals(array2d) As Boolean
    Equals = False
    If lb(1) <> LBound(array2d, 1) Then Exit Function
    If lb(2) <> LBound(array2d, 2) Then Exit Function
    If ub(1) <> UBound(array2d, 1) Then Exit Function
    If ub(2) <> UBound(array2d, 2) Then Exit Function
    Dim i As Long, j As Long
    For i = lb(1) To ub(1)
        For j = lb(2) To ub(2)
            If array_(i, j) <> array2d(i, j) Then Exit Function
        Next j
    Next i
    Equals = True
End Function

Public Function ShiftIndex(ri As Long, ci As Long) As ArrayEx2
    Dim re, rearr As New ArrayEx2
    
    re = ArrayExCore.ShiftIndex(array_, ri, ci)
    
    Set ShiftIndex = rearr.Init(re)
End Function

' Set
Public Function SetElements(argRows As Variant, argCols As Variant, argVal As Variant) As ArrayEx2
    Dim rows, cols
    rows = ArrayExCore.IndexToArray(argRows, lb(1), ub(1))
    cols = ArrayExCore.IndexToArray(argCols, lb(2), ub(2))
    If UBound(rows) <> UBound(argVal, 1) Then Call Err.Raise(2102, "SetElements", "ArrayEx: Need same size")
    If UBound(cols) <> UBound(argVal, 2) Then Call Err.Raise(2102, "SetElements", "ArrayEx: Need same size")

    Dim i As Long, j As Long, ii As Long, jj As Long
    Dim re
    ii = 1
    For i = LBound(rows) To UBound(rows)
        jj = 1
        For j = LBound(cols) To UBound(cols)
            array_(rows(i), cols(j)) = argVal(ii, jj)
            jj = jj + 1
        Next j
        ii = ii + 1
    Next i
    Set SetElements = Me
End Function

Public Function SetElement(argRows As Long, argCols As Long, argVal As Variant) As ArrayEx2
    array_(argRows, argCols) = argVal
    Set SetElement = Me
End Function

Public Function SetRow(argRow As Long, argArrayEx1 As ArrayEx1)
    Dim i
    If lb(2) <> argArrayEx1.lb And ub(2) <> argArrayEx1.ub Then _
            Call Err.Raise(2102, "SetRow", "ArrayEx: Need same length")
    
    Dim re:  re = argArrayEx1.Value
    For i = lb(2) To ub(2)
        array_(argRow, i) = re(i)
    Next i
    Set SetRow = Me
End Function

Public Function SetColumn(argCol As Long, argArrayEx1 As ArrayEx1)
    Dim i
    If lb(1) <> argArrayEx1.lb And ub(1) <> argArrayEx1.ub Then _
            Call Err.Raise(2102, "SetColumn", "ArrayEx: Need same length")
    
    Dim re:  re = argArrayEx1.Value
    For i = lb(1) To ub(1)
        array_(i, argCol) = re(i)
    Next i
    Set SetColumn = Me
End Function

Public Function Extract(rows As Variant, cols As Variant) As ArrayEx2
    Set Extract = ExtractCore(rows, cols)
End Function

Public Function GetRows(rows As Variant) As ArrayEx2
    Set GetRows = ExtractCore(rows, ":")
End Function

Public Function GetColumns(cols As Variant) As ArrayEx2
    Set GetColumns = ExtractCore(":", cols)
End Function

Function GetRow(row As Long) As ArrayEx1
    Dim re, rearr As New ArrayEx1
    
    re = ArrayExCore.ConvertToDimensionN(ExtractCore(row, ":").Value, 1)
    
    Set GetRow = rearr.Init(re)
End Function

Public Function GetColumn(col As Long) As ArrayEx1
    Dim re, rearr As New ArrayEx1
    
    re = ArrayExCore.ConvertToDimensionN(ExtractCore(":", col).Transpose.Value, 1)
    
    Set GetColumn = rearr.Init(re)
End Function

Public Function GetElement(row As Long, col As Long) As ArrayEx0
    Dim rre:   rre = ExtractCore(row, col).Value
    Dim rearr As New ArrayEx0
    Set GetElement = rearr.Init(rre(LBound(rre, 1), LBound(rre, 2)))
End Function

' Add
Function AddRow(rowArrayEx1 As ArrayEx1) As ArrayEx2
    Dim rearr As New ArrayEx2
    Set AddRow = rearr.InitForce(ArrayExCore.VSTACK(array_, rowArrayEx1.Value))
End Function

Function AddColumn(colArrayEx1 As ArrayEx1) As ArrayEx2
    Dim rearr As New ArrayEx2
    Set AddColumn = rearr.InitForce(ArrayExCore.HSTACK(array_, colArrayEx1.Value))
End Function

' Convert
Public Function ToString() As String
    Dim s As String, i As Long, j As Long
    s = "["
    For i = lb(1) To ub(1)
        For j = lb(2) To ub(2)
            s = s & CStr(array_(i, j)) & ","
        Next j
        s = Left(s, Len(s) - 1) & ";"
    Next i
    s = Left(s, Len(s) - 1) & "]"
    ToString = s
End Function

Public Function ToRange(rng As Range)
    rng.Resize(Count(1), Count(2)).Value = array_
End Function

Public Function ToCollection() As Collection
    Dim coll As New Collection
    Dim arr As New ArrayEx1
    Dim i As Long
    For i = lb To ub
        Set arr = GetRow(i)
        coll.Add arr
    Next i
    Set ToCollection = coll
End Function

' Linqs
Public Function First() As ArrayEx1
    Set First = GetRow(lb)
End Function

Public Function Last() As ArrayEx1
    Set Last = GetRow(ub)
End Function

Public Function Skip(skipIndex As Long) As ArrayEx2
    Set Skip = ExtractCore(skipIndex & ":" & ub, ":")
End Function

Public Function TAKE(skipIndex As Long) As ArrayEx2
    Set TAKE = ExtractCore(lb & ":" & skipIndex - 1, ":")
End Function

Public Function Distinct(column As Long) As ArrayEx2
    Dim arr As New ArrayEx1
    Set arr = GetColumn(column)
    Dim i As Long, s As String, sIndexes As String
    For i = arr.lb To arr.ub
        If InStr(s, arr.GetElement(i)) = 0 Then
            s = s & arr.GetElement(i) & ","
            sIndexes = sIndexes & i & ","
        End If
    Next i
    If sIndexes = "" Then Exit Function
    sIndexes = Left(sIndexes, Len(sIndexes) - 1)
    Set Distinct = ExtractCore(sIndexes, ":")
End Function

Public Function OrderBy(column As Long) As ArrayEx2
    Dim re As New ArrayEx2
    Set OrderBy = re.Init(WorksheetFunction.Sort(array_, column))
End Function

Public Function OrderByDescending(column As Long) As ArrayEx2
    Dim re As New ArrayEx2
    Set OrderByDescending = re.Init(WorksheetFunction.Sort(array_, column, -1, False))
End Function

Public Function Reverse() As ArrayEx2
    Dim rearr: ReDim rearr(lb To ub)
    Dim i As Long
    For i = lb To ub
        rearr(i) = array_(ub - i + 1)
    Next i
    Dim re As New ArrayEx2
    Set Reverse = ExtractCore(rearr, ":")
End Function

Public Function AllEvaluate(arguments, columnNumbers, formula As String) As Boolean
    AllEvaluate = False
    Dim evaluation
    For Each evaluation In GetEvaluations(arguments, columnNumbers, formula)
        If evaluation = False Then Exit Function
    Next
    AllEvaluate = True
End Function

Public Function AnyEvaluate(arguments, columnNumbers, formula As String) As Boolean
    AnyEvaluate = True
    Dim evaluation
    For Each evaluation In GetEvaluations(arguments, columnNumbers, formula)
        If evaluation = True Then Exit Function
    Next
    AnyEvaluate = False
End Function

Public Function WhereEvaluated(arguments, columnNumbers, formula As String) As ArrayEx2
    Dim evaluations, i As Long, indexes As String
    evaluations = GetEvaluations(arguments, columnNumbers, formula)
    For i = lb To ub
        If evaluations(i) = True Then indexes = indexes & i & ","
    Next
    Set WhereEvaluated = ExtractCore(indexes, ":")
End Function

Public Function SelectEvaluated(arguments, columnNumbers, formula As String) As ArrayEx1
    Dim evaluations, i As Long, j As Long, s As String, indexes As String
    evaluations = GetEvaluations(arguments, columnNumbers, formula)
    Dim re As New ArrayEx1
    Set SelectEvaluated = re.Init(evaluations)
End Function

' DebugPrint
Public Function DebugPrint(argWithCommna As String, ColumnNumbersWithConnma As String, formulaWithArg As String) As ArrayEx2
    Dim args: args = Split(Replace(argWithCommna, " ", ""), ",")
    Dim cols: cols = Split(Replace(ColumnNumbersWithConnma, " ", ""), ",")
    Dim i&, j&, s$
    For i = lb To ub
        s = formulaWithArg
        For j = LBound(args) To UBound(args)
            s = Replace(s, args(j), array_(i, val(cols(j))))
        Next j
        Debug.Print s
    Next i
    Set DebugPrint = Me
End Function

Public Function DebugPrintAll() As ArrayEx2
    Dim s: s = ToString()
    Dim ss:  ss = Replace(s, ";", ";" & vbCrLf)
    Debug.Print ss
    Set DebugPrintAll = Me
End Function

' Wrap default functions
Public Function Transpose() As ArrayEx2
    Dim rearr As New ArrayEx2
    Set Transpose = rearr.InitForce(WorksheetFunction.Transpose(array_))
End Function

Public Function VSTACK(arr, ParamArray paramarr()) As ArrayEx2
    Dim aex2 As New ArrayEx2
    Set VSTACK = aex2.Init(ArrayExCore.VSTACK(array_, arr, paramarr))
End Function





' Private Methods --------------------------------------------------------------------------------------------
Private Function ExtractCore(argRows As Variant, argCols As Variant) As ArrayEx2
    Dim re, rearr As New ArrayEx2
    
    re = ArrayExCore.ExtractCore(array_, argRows, argCols)
    
    Set ExtractCore = rearr.Init(re)
End Function

Private Function GetEvaluations(arguments, columnNumbers, formula As String) As Variant
    Dim i As Long, j As Long, s As String, indexes As String, re
    Dim args, cols, rearr
    
    ReDim rearr(lb To ub)
    args = Split(arguments, ",")
    cols = Split(columnNumbers, ",")
    If UBound(args) <> UBound(cols) Then Err.Raise 2002, "ArrayEx2, GetEvaluations", "arguments and columnNumbers count are not same."
    
    For i = lb To ub
        s = formula
        For j = LBound(args) To UBound(args)
            s = Replace(s, args(j), array_(i, cols(j)))
        Next j
        If ArrayExCore.TryEvaluate(s, re) Then rearr(i) = re
    Next i
    GetEvaluations = rearr
End Function

