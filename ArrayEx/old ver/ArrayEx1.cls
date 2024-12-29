VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayEx1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'<dir .\old ver /dir>
Option Explicit

' ------------------------------------------------------------------------------------------------------------
' Project                   : https://github.com/yyukki5/ArrayEx
' Version                   : 2.1.0
' Dependencies              : ArrayExCore.cls, ArrayEx0.cls
' ------------------------------------------------------------------------------------------------------------

' Fields -----------------------------------------------------------------------------------------------------
Private array_ As Variant


' Properties -------------------------------------------------------------------------------------------------
Public Property Get Value(): Value = array_: End Property
Public Property Let Value(array1d): Init array1d: End Property
Public Property Get ub(): ub = UBound(array_, 1): End Property
Public Property Get lb(): lb = LBound(array_, 1): End Property
Public Property Get Count() As Long: Count = UBound(array_, 1) - LBound(array_, 1) + 1: End Property


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
    If lb <> LBound(array1d) Then Exit Function
    If ub <> UBound(array1d) Then Exit Function
    Dim i As Long
    For i = lb To ub
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
Public Function AddElement(element As ArrayEx0) As ArrayEx1
    If IsEmpty(array_) Then
        array_ = Array(element.Value)
    Else
        ReDim Preserve array_(lb To ub + 1): array_(ub) = element.Value
    End If
    Set AddElement = Me
End Function

' Convert
Public Function ToString() As String
    Dim s As String, i As Long, j As Long
    s = "{"
    For i = lb To ub
        s = s & CStr(array_(i)) & ","
    Next i
    s = Left(s, Len(s) - 1) & "}"
    ToString = s
End Function

Public Function ToCollection() As Collection
    Dim coll As New Collection
    Dim arr As ArrayEx0
    Dim i As Long
    
    For i = lb To ub
        Set arr = New ArrayEx0
        coll.Add arr.Init(array_(i))
    Next i

    Set ToCollection = coll
End Function

Public Function SetExcelRange(rng As Object) As Object
    rng.Resize(1, Count).Value = array_
    Set SetExcelRange = rng.Resize(1, Count)
End Function

' linqs
Public Function Contains(element As Variant) As Boolean
    Contains = True
    Dim i As Long
    For i = lb To ub
        If array_(i) = element Then Exit Function
    Next i
    Contains = False
End Function

Public Function First() As ArrayEx0
    Set First = GetElement(lb)
End Function

Public Function Last() As ArrayEx0
    Set Last = GetElement(ub)
End Function

Public Function Skip(skipIndex As Long) As ArrayEx1
    Set Skip = ExtractCore(lb + skipIndex & ":" & ub)
End Function

Public Function Take(takeIndex As Long) As ArrayEx1
    Set Take = ExtractCore(lb & ":" & takeIndex + lb - 1)
End Function

Public Function Distinct() As ArrayEx1
    Dim var As Variant
    Dim i As Long, s As String, sIndexes As String
    Dim coll As New Collection
    For i = lb To ub
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
    Set OrderByAscending = re.Init(Sort(array_, True))
End Function

Public Function OrderByDescending() As ArrayEx1
    Dim re As New ArrayEx1
    Set OrderByDescending = re.Init(Sort(array_, False))
End Function

Public Function Reverse() As ArrayEx1
    Dim rearr: ReDim rearr(lb To ub)
    Dim i As Long
    For i = lb To ub
        rearr(i) = array_(ub - i + 1)
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
    For i = lb To ub
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

Public Function DebugPrint(Optional formatString As String = "") As ArrayEx1
    Dim i&, s$
    For i = lb To ub
        Debug.Print Format(array_(i), formatString)
    Next i
    Set DebugPrint = Me
End Function

Public Function Match(arg, Optional matchType As Long = 0) As ArrayEx0
    Dim re As New ArrayEx0
    Set Match = re.Init(ArrayMatch(arg, array_, matchType))
End Function

Public Function Max() As ArrayEx0
    Dim re As New ArrayEx0, val, temp
    temp = array_(LBound(array_))
    For Each val In array_
        If val > temp Then temp = val
    Next
    Set Max = re.Init(temp)
End Function
Public Function Min() As ArrayEx0
    Dim re As New ArrayEx0, val, temp
    temp = array_(LBound(array_))
    For Each val In array_
        If val < temp Then temp = val
    Next
    Set Min = re.Init(temp)
End Function
Public Function Sum() As ArrayEx0
    Dim re As New ArrayEx0, val, temp
    For Each val In array_
        temp = temp + val
    Next
    Set Sum = re.Init(temp)
End Function
Public Function Average() As ArrayEx0
    Dim re As New ArrayEx0
    Set Average = re.Init((Sum.Value / Count))
End Function

' [Hack] Wrap default functions. Better to use app.*** in case of depending on application function.


' Private Methods ------------------------------------------------------------------
Private Function ExtractCore(argIndexs As Variant) As ArrayEx1
    Dim re, rearr As New ArrayEx1
    re = ArrayExCore.ExtractCore(array_, argIndexs)
    Set ExtractCore = rearr.Init(re)
End Function

Private Function GetEvaluations(argument As String, formula As String) As Variant
    Dim i As Long, s As String, rearr, re
    ReDim rearr(lb To ub)
    For i = lb To ub
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

Private Function ArrayMatch(val As Variant, arr As Variant, matchType As Long) As Variant
    Dim v, i As Long
    
    For i = LBound(arr) To UBound(arr)
        v = arr(i)
        Select Case matchType
            Case 1: If v > val Then ArrayMatch = i - 1: Exit For
            Case 0: If v = val Then ArrayMatch = i: Exit For
            Case -1: If v < val Then ArrayMatch = i - 1: Exit For
            Case Else:      Err.Raise 9101, "ArrayMatch", "undefined matchType value"
        End Select
    Next i
    If ArrayMatch < LBound(arr) Or ArrayMatch > UBound(arr) Then Err.Raise 9101, "ArrayMatch", "Can not find value"
End Function

Private Function Sort(ByVal arr, isAsc As Boolean) As Variant
    QuickSort arr, LBound(arr), UBound(arr), isAsc
    Sort = arr
End Function

Private Function QuickSort(ByRef arr As Variant, ByVal low As Long, ByVal up As Long, isAsc As Boolean)
    Dim i As Long, j As Long, pivot As Variant, temp As Variant
    i = low
    j = up
    pivot = arr((low + up) \ 2)
    
    Do While i <= j
        If isAsc Then
            Do While arr(i) < pivot: i = i + 1: Loop
            Do While arr(j) > pivot: j = j - 1: Loop
        Else
            Do While arr(i) > pivot: i = i + 1: Loop
            Do While arr(j) < pivot: j = j - 1: Loop
        End If
        If i <= j Then
            temp = arr(i)
            arr(i) = arr(j)
            arr(j) = temp
            i = i + 1
            j = j - 1
        End If
    Loop
    If low < j Then QuickSort arr, low, j, isAsc
    If i < up Then QuickSort arr, i, up, isAsc
    QuickSort = arr
End Function


