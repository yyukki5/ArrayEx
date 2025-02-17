VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayEx2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'<dir .\old ver /dir>
Option Explicit

' ------------------------------------------------------------------------------------------------------------
' Project                   : https://github.com/yyukki5/ArrayEx
' Version                   : 2.1.0
' Dependencies              : ArrayExCore.cls, ArrayEx0.cls, ArrayEx1.cls
' ------------------------------------------------------------------------------------------------------------

' Fields -----------------------------------------------------------------------------------------------------
Private array_ As Variant


' Properties -------------------------------------------------------------------------------------------------
Public Property Get Value(): Value = array_: End Property
Public Property Let Value(array2d): Init array2d: End Property
Public Property Get lb1() As Long: lb1 = LBound(array_, 1): End Property
Public Property Get lb2() As Long: lb2 = LBound(array_, 2): End Property
Public Property Get ub1() As Long: ub1 = UBound(array_, 1): End Property
Public Property Get ub2() As Long: ub2 = UBound(array_, 2): End Property
Public Property Get Count1() As Long: Count1 = UBound(array_, 1) - LBound(array_, 1) + 1: End Property
Public Property Get Count2() As Long: Count2 = UBound(array_, 2) - LBound(array_, 2) + 1: End Property
Public Property Get Rank() As Long: Rank = 2: End Property

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

Public Function Equals(arrayEx2d As ArrayEx2) As Boolean
    Dim array2d: array2d = arrayEx2d.Value
    
    Equals = False
    If lb1 <> LBound(array2d, 1) Then Exit Function
    If lb2 <> LBound(array2d, 2) Then Exit Function
    If ub1 <> UBound(array2d, 1) Then Exit Function
    If ub2 <> UBound(array2d, 2) Then Exit Function
    Dim i As Long, j As Long
    For i = lb1 To ub1
        For j = lb2 To ub2
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
Public Function SetElement(rowIndexes As Variant, columnIndexes As Variant, val As Variant) As ArrayEx2
    Dim rIndexes, cIndexes
    Dim i As Long, j As Long

    rIndexes = ArrayExCore.IndexToArray(rowIndexes, lb1, ub1)
    cIndexes = ArrayExCore.IndexToArray(columnIndexes, lb2, ub2)
    Dim re
    For i = LBound(rIndexes) To UBound(rIndexes)
        For j = LBound(cIndexes) To UBound(cIndexes)
            array_(rIndexes(i), cIndexes(j)) = val
        Next j
    Next i
    Set SetElement = Me
End Function

Public Function SetRow(rowIndexes As Long, argArrayEx1 As ArrayEx1)
    Dim i
    If lb2 <> argArrayEx1.lb And ub2 <> argArrayEx1.ub Then _
            Call Err.Raise(2102, "SetRow", "ArrayEx: Need same length")
    
    Dim re:  re = argArrayEx1.Value
    For i = lb2 To ub2
        array_(rowIndexes, i) = re(i)
    Next i
    Set SetRow = Me
End Function

Public Function SetColumn(columnIndexes As Long, argArrayEx1 As ArrayEx1)
    Dim i
    If lb1 <> argArrayEx1.lb And ub1 <> argArrayEx1.ub Then _
            Call Err.Raise(2102, "SetColumn", "ArrayEx: Need same length")
    
    Dim re:  re = argArrayEx1.Value
    For i = lb1 To ub1
        array_(i, columnIndexes) = re(i)
    Next i
    Set SetColumn = Me
End Function

Public Function Extract(rowIndexes As Variant, columnIndexes As Variant) As ArrayEx2
    Set Extract = ExtractCore(rowIndexes, columnIndexes)
End Function

Public Function GetRows(rowIndexes As Variant) As ArrayEx2
    Set GetRows = ExtractCore(rowIndexes, ":")
End Function

Public Function GetColumns(columnIndexes As Variant) As ArrayEx2
    Set GetColumns = ExtractCore(":", columnIndexes)
End Function

Function GetRow(rowIndex As Long) As ArrayEx1
    Dim re, rearr As New ArrayEx1

    re = ArrayExCore.ConvertToDimensionN(ExtractCore(rowIndex, ":").Value, 1)
    
    Set GetRow = rearr.Init(re)
End Function

Public Function GetColumn(columnIndex As Long) As ArrayEx1
    Dim re, rearr As New ArrayEx1
    
    re = ArrayExCore.ConvertToDimensionN(ExtractCore(":", columnIndex).Transpose.Value, 1)
    
    Set GetColumn = rearr.Init(re)
End Function

Public Function GetElement(rowIndex As Long, columnIndex As Long) As ArrayEx0
    Dim rre:   rre = ExtractCore(rowIndex, columnIndex).Value
    Dim rearr As New ArrayEx0
    Set GetElement = rearr.Init(rre(LBound(rre, 1), LBound(rre, 2)))
End Function

' Add
Function AddRow(rowArrayEx1 As ArrayEx1) As ArrayEx2
    If IsEmpty(array_) Then Set AddRow = Me.InitForce(rowArrayEx1.Value) Else Set AddRow = Me.InitForce(ArrayExCore.VSTACK(array_, rowArrayEx1.Value))
End Function

Function AddColumn(colArrayEx1 As ArrayEx1) As ArrayEx2
    Dim ae2 As New ArrayEx2
    If IsEmpty(array_) Then Set AddColumn = Me.InitForce(colArrayEx1.Value).Transpose Else Set AddColumn = Me.InitForce(ArrayExCore.HSTACK(array_, ae2.InitForce(colArrayEx1.Value).Transpose.Value))
End Function

' Convert
Public Function ToString() As String
    Dim s As String, i As Long, j As Long
    s = "{"
    For i = lb1 To ub1
        For j = lb2 To ub2
            s = s & CStr(array_(i, j)) & ","
        Next j
        s = Left(s, Len(s) - 1) & ";"
    Next i
    s = Left(s, Len(s) - 1) & "}"
    ToString = s
End Function

Public Function ToCollection(Optional dimension As Long = 0) As Collection
    Dim coll As New Collection
    Dim arr0 As ArrayEx0, arr1 As ArrayEx1
    Dim i As Long, ar
    Select Case dimension
        Case 1
            For i = lb1 To ub1
                Set arr1 = New ArrayEx1
                Set arr1 = GetRow(i)
                coll.Add arr1
            Next i
        Case 2
            For i = lb2 To ub2
                Set arr1 = New ArrayEx1
                Set arr1 = GetColumn(i)
                coll.Add arr1
            Next i
        Case Else
            For Each ar In array_
                Set arr0 = New ArrayEx0
                coll.Add arr0.Init(ar)
            Next
    End Select
    Set ToCollection = coll
End Function

Public Function SetExcelRange(rng As Object) As Object
    rng.Resize(Count1, Count2).Value = array_
    Set SetExcelRange = rng.Resize(Count1, Count2)
End Function

' Linqs
Public Function First() As ArrayEx1
    Set First = GetRow(lb1)
End Function

Public Function Last() As ArrayEx1
    Set Last = GetRow(ub1)
End Function

Public Function Skip(skipIndex As Long) As ArrayEx2
    Set Skip = ExtractCore(lb1 + skipIndex & ":" & ub1, ":")
End Function

Public Function Take(skipIndex As Long) As ArrayEx2
    Set Take = ExtractCore(lb1 & ":" & skipIndex + lb1 - 1, ":")
End Function

Public Function Distinct(columnIndex As Long) As ArrayEx2
    Dim i As Long, s As String, sIndexes As String
    
    Dim arrayCollection As Collection
    Set arrayCollection = ToCollection(1)
    Dim coll As New Collection
    For i = lb1 To ub1
        If IsDuplicated(coll, array_(i, columnIndex)) Then
        Else
            coll.Add array_(i, columnIndex)
            sIndexes = sIndexes & i & ","
        End If
    Next i
    
    If sIndexes = "" Then Exit Function
    sIndexes = Left(sIndexes, Len(sIndexes) - 1)
    Set Distinct = ExtractCore(sIndexes, ":")
End Function

Public Function OrderByAscending(columnIndex As Long) As ArrayEx2
    Dim re As New ArrayEx2:    Set OrderByAscending = re.Init(Sort(array_, columnIndex, True))
End Function

Public Function OrderByDescending(columnIndex As Long) As ArrayEx2
    Dim re As New ArrayEx2:    Set OrderByDescending = re.Init(Sort(array_, columnIndex, False))
End Function

Public Function Reverse() As ArrayEx2
    Dim rearr: ReDim rearr(lb1 To ub1)
    Dim i As Long
    For i = lb1 To ub1
        rearr(i) = ub1 + 1 - i
    Next i
    Dim re As New ArrayEx2
    Set Reverse = ExtractCore(rearr, ":")
End Function

Public Function AllEvaluate(arguments, columnIndexes, formula As String) As Boolean
    AllEvaluate = False
    Dim evaluation
    For Each evaluation In GetEvaluations(arguments, columnIndexes, formula)
        If VarType(evaluation) <> vbBoolean Then Err.Raise 2101, "ArrayEx2.AllEvaluate()", "Should evaluated value boolean, but not boolean: " & evaluation
        If evaluation = False Then Exit Function
    Next
    AllEvaluate = True
End Function

Public Function AnyEvaluate(arguments, columnIndexes, formula As String) As Boolean
    AnyEvaluate = True
    Dim evaluation
    For Each evaluation In GetEvaluations(arguments, columnIndexes, formula)
        If evaluation = True Then Exit Function
    Next
    AnyEvaluate = False
End Function

Public Function WhereEvaluated(arguments, columnIndexes, formula As String) As ArrayEx2
    Dim evaluations, i As Long, indexes As String
    evaluations = GetEvaluations(arguments, columnIndexes, formula)
    For i = lb1 To ub1
        If evaluations(i) = True Then indexes = indexes & i & ","
    Next
    Set WhereEvaluated = ExtractCore(indexes, ":")
End Function

Public Function SelectEvaluated(arguments, columnIndexes, formula As String) As ArrayEx1
    Dim evaluations, i As Long, j As Long, s As String, indexes As String
    evaluations = GetEvaluations(arguments, columnIndexes, formula)
    Dim re As New ArrayEx1
    Set SelectEvaluated = re.Init(evaluations)
End Function

Public Function EachEvaluated(argument, formula As String) As ArrayEx2
    Dim i As Long, j As Long, re As ArrayEx2, s, val, rearr: rearr = array_
    For i = lb1 To ub1
        For j = lb2 To ub2
            If ArrayExCore.TryEvaluate(Replace(formula, argument, rearr(i, j)), val) Then rearr(i, j) = val
        Next j
    Next i
    Set EachEvaluated = re.Init(rearr)
End Function

' DebugPrint
Public Function DebugPrint(argWithCommna As String, columnIndexesWithConnma As String, formulaWithArg As String) As ArrayEx2
    Dim args: args = Split(Replace(argWithCommna, " ", ""), ",")
    Dim cols: cols = Split(Replace(columnIndexesWithConnma, " ", ""), ",")
    Dim i&, j&, s$
    For i = lb1 To ub1
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

Public Function Transpose() As ArrayEx2
    Dim re, rearr As New ArrayEx2, i As Long, j As Long
    
    ReDim re(lb2 To ub2, lb1 To ub1)
    For i = lb2 To ub2: For j = lb1 To ub1: re(i, j) = array_(j, i): Next j: Next i
    
    Set Transpose = rearr.Init(re)
End Function
Public Function VSTACK(arr, ParamArray paramarr()) As ArrayEx2
    Dim aex2 As New ArrayEx2:    Set VSTACK = aex2.Init(ArrayExCore.VSTACK(array_, arr, paramarr))
End Function

Public Function HSTACK(arr, ParamArray paramarr()) As ArrayEx2
    Dim aex2 As New ArrayEx2:    Set HSTACK = aex2.Init(ArrayExCore.HSTACK(array_, arr, paramarr))
End Function

' [Hack] Wrap default functions. Better to use app.*** in case of depending on application function.


' Private Methods --------------------------------------------------------------------------------------------
Private Function ExtractCore(rowIndexes As Variant, columnIndexes As Variant) As ArrayEx2
    Dim re, rearr As New ArrayEx2
    
    re = ArrayExCore.ExtractCore(array_, rowIndexes, columnIndexes)
    
    Set ExtractCore = rearr.Init(re)
End Function

Private Function GetEvaluations(arguments, columnIndexes, formula As String) As Variant
    Dim i As Long, j As Long, s As String, indexes As String, re
    Dim args, cols, rearr
    
    ReDim rearr(lb1 To ub1)
    args = Split(arguments, ",")
    cols = Split(columnIndexes, ",")
    If cols(LBound(cols)) < lb2 Then Err.Raise 2002, "ArrayEx2.GetEvaluations()", "Column number is out of array index (" & cols(LBound(cols)) & ")."
    If cols(UBound(cols)) > ub2 Then Err.Raise 2002, "ArrayEx2.GetEvaluations()", "Column number is out of array index (" & cols(UBound(cols)) & ")."
    If UBound(args) <> UBound(cols) Then Err.Raise 2002, "ArrayEx2.GetEvaluations()", "Number of arguments and columnIndexes are not same."
    
    For i = lb1 To ub1
        s = formula
        For j = LBound(args) To UBound(args)
            s = Replace(s, args(j), array_(i, cols(j)))
        Next j
        If ArrayExCore.TryEvaluate(s, re) Then rearr(i) = re Else Err.Raise 2003, "ArrayEx2.GetEvaluations()", "String in evaluate() is error :" & s & "."
    Next i
    GetEvaluations = rearr
End Function

Private Function IsDuplicated(coll As Collection, val) As Boolean
    Dim col
    On Error GoTo errEqual
    IsDuplicated = True
    For Each col In coll
        If col = val Then Exit Function
    Next
    IsDuplicated = False
errEqual:
End Function

Private Function Sort(ByVal arr, col As Long, isAsc As Boolean) As Variant
    QuickSort arr, LBound(arr, 1), UBound(arr, 1), isAsc, col
    Sort = arr
End Function

Private Function QuickSort(ByRef arr As Variant, ByVal low As Long, ByVal up As Long, isAsc As Boolean, col As Long)
    Dim i As Long, j As Long, k As Long, pivot As Variant, temp As Variant
    i = low
    j = up
    pivot = arr((low + up) \ 2, col)
    
    Do While i <= j
        If isAsc Then
            Do While arr(i, col) < pivot: i = i + 1: Loop
            Do While arr(j, col) > pivot: j = j - 1: Loop
        Else
            Do While arr(i, col) > pivot: i = i + 1: Loop
            Do While arr(j, col) < pivot: j = j - 1: Loop
        End If
        If i <= j Then
            For k = LBound(arr, 2) To UBound(arr, 2)
                temp = arr(i, k)
                arr(i, k) = arr(j, k)
                arr(j, k) = temp
            Next k
            i = i + 1
            j = j - 1
        End If
    Loop
    If low < j Then QuickSort arr, low, j, isAsc, col
    If i < up Then QuickSort arr, i, up, isAsc, col
    QuickSort = arr
End Function

