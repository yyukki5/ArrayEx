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
Public Property Get Value(Optional rows As Variant, Optional cols As Variant)
Attribute Value.VB_UserMemId = 0
    If IsMissing(rows) = True And IsMissing(cols) = True Then Value = array_
    If IsMissing(rows) = False And IsMissing(cols) = True Then Value = ExtractCore(rows, ":").ToPrimitive
    If IsMissing(rows) = True And IsMissing(cols) = False Then Value = ExtractCore(":", cols).ToPrimitive
    If IsMissing(rows) = False And IsMissing(cols) = False Then Value = ExtractCore(rows, cols).ToPrimitive
End Property

' [Hack] need to more good idea.
' for argument.  need for return value as ArrayEx2.
Public Property Get Mine()
    Set Mine = Me
End Property

Public Property Get Ub(Optional direction As Long = 1) As Long
    Ub = UBound(array_, direction)
End Property

Public Property Get Lb(Optional direction As Long = 1) As Long
    Lb = LBound(array_, direction)
End Property

Public Property Get Count() As Long
    Count = UBound(array_, 1) - LBound(array_, 1) + 1
End Property

' Public Methods ---------------------------------------------------------------------------------------------
Public Function Init(array2d As Variant) As ArrayEx2
    If IsArrayRank(array2d, 2) = False Then Call Err.Raise(2101, "ArrayEx2 Init", "ArrayEx: Need 2d array")
    array_ = array2d
    Set Init = Me
End Function

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
    If Lb(1) <> LBound(array2d, 1) Then Exit Function
    If Lb(2) <> LBound(array2d, 2) Then Exit Function
    If Ub(1) <> UBound(array2d, 1) Then Exit Function
    If Ub(2) <> UBound(array2d, 2) Then Exit Function
    Dim i As Long, j As Long
    For i = Lb(1) To Ub(1)
        For j = Lb(2) To Ub(2)
            If array_(i, j) <> array2d(i, j) Then Exit Function
        Next j
    Next i
    Equals = True
End Function

Public Function ToPrimitive() As Variant
    If Ub(1) = 1 And Ub(2) = 1 Then ToPrimitive = array_(1, 1)
    If Ub(1) = 1 And Ub(2) <> 1 Then ToPrimitive = GetRow(Ub(1)).Value()
    If Ub(1) <> 1 And Ub(2) = 1 Then ToPrimitive = GetColumn(Ub(2)).Value()
    If Ub(1) <> 1 And Ub(2) <> 1 Then ToPrimitive = array_
End Function

' Set
Public Function SetElements(argRows As Variant, argCols As Variant, argVal As Variant) As ArrayEx2
    Dim rows, cols
    rows = argToArray(argRows, 1)
    cols = argToArray(argCols, 2)
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
    If Lb(2) <> argArrayEx1.Lb And Ub(2) <> argArrayEx1.Ub Then _
                                         Call Err.Raise(2102, "SetRow", "ArrayEx: Need same length")
    
    Dim re:  re = argArrayEx1.Value
    For i = Lb(2) To Ub(2)
        array_(argRow, i) = re(i)
    Next i
    Set SetRow = Me
End Function

Public Function SetColumn(argCol As Long, argArrayEx1 As ArrayEx1)
    Dim i
    If Lb(1) <> argArrayEx1.Lb And Ub(1) <> argArrayEx1.Ub Then _
                                         Call Err.Raise(2102, "SetColumn", "ArrayEx: Need same length")
    
    Dim re:  re = argArrayEx1.Value
    For i = Lb(1) To Ub(1)
        array_(i, argCol) = re(i)
    Next i
    Set SetColumn = Me
End Function

' Get
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
    Dim re: re = ExtractCore(row, ":").Value
    Dim rre: ReDim rre(1 To UBound(re, 2) - LBound(re, 2) + 1)
    Dim i As Long: For i = 1 To UBound(re, 2) - LBound(re, 2) + 1: rre(i) = re(1, i): Next i
    Dim rearr As New ArrayEx1
    Set GetRow = rearr.Init(rre)
End Function

Public Function GetColumn(col As Long) As ArrayEx1
    Dim re: re = ExtractCore(":", col).Value
    Dim rre: ReDim rre(1 To UBound(re, 1) - LBound(re, 1) + 1)
    Dim i As Long: For i = 1 To UBound(re, 1) - LBound(re, 1) + 1: rre(i) = re(i, 1): Next i
    Dim rearr As New ArrayEx1
    Set GetColumn = rearr.Init(rre)
End Function

Public Function GetElement(row As Long, col As Long) As ArrayEx0
    Dim rre:   rre = ExtractCore(row, col).Value(1, 1)
    Dim rearr As New ArrayEx0
    Set GetElement = rearr.Init(rre)
End Function

' Add
Function AddRow(row As ArrayEx1) As ArrayEx2
    Dim i, j
    Dim re
    If IsEmpty(array_) Then
        ReDim re(1 To 1, row.Lb To row.Ub)
        For j = LBound(re, 2) To UBound(re, 2)
            re(1, j) = row(j)
        Next j
    Else
        Dim maxCol As Long
        maxCol = IIf(Ub(2) > row.Ub, Ub(2), row.Ub)
        
        ReDim re(Lb To Ub + 1, Lb(2) To maxCol)
        For i = Lb To Ub
            For j = Lb(2) To Ub(2)
                re(i, j) = array_(i, j)
            Next j
        Next i
        For j = Lb(2) To maxCol
            re(Ub + 1, j) = row(j)
        Next j
    End If
    Dim rearr As New ArrayEx2
    Set AddRow = rearr.Init(re)
End Function

Function AddColumn(col As ArrayEx1) As ArrayEx2
    Dim i, j
    Dim re
    If IsEmpty(array_) Then
        ReDim re(col.Lb To col.Ub, 1 To 1)
        For j = LBound(re, 2) To UBound(re, 2)
            re(j, 1) = col(j)
        Next j
    Else
        Dim maxRow As Long
        maxRow = IIf(Ub(1) > col.Ub, Ub(1), col.Ub)
        
        ReDim re(Lb To maxRow, Lb(2) To Ub(2) + 1)
        For i = Lb To Ub
            For j = Lb(2) To Ub(2)
                re(i, j) = array_(i, j)
            Next j
        Next i
        For j = Lb(1) To maxRow
            re(j, Ub(2) + 1) = col(j)
        Next j
    End If
    Dim rearr As New ArrayEx2
    Set AddColumn = rearr.Init(re)
End Function

' Convert
Public Function ToString() As String
    Dim s As String, i As Long, j As Long
    s = "["
    For i = Lb(1) To Ub(1)
        For j = Lb(2) To Ub(2)
            s = s & CStr(array_(i, j)) & ","
        Next j
        s = Left(s, Len(s) - 1) & ";"
    Next i
    s = Left(s, Len(s) - 1) & "]"
    ToString = s
End Function

Public Function ToRange(rng As Range)
    rng.Resize(Ub(1), Ub(2)).Value = array_
End Function

Public Function ToCollection() As Collection
    Dim coll As New Collection
    Dim arr As New ArrayEx1
    Dim i As Long
    For i = Lb To Ub
        Set arr = GetRow(i)
        coll.Add arr
    Next i
    Set ToCollection = coll
End Function

' Linqs
Public Function First() As ArrayEx1
    Set First = GetRow(Lb)
End Function

Public Function Last() As ArrayEx1
    Set Last = GetRow(Ub)
End Function

Public Function Skip(skipIndex As Long) As ArrayEx2
    Set Skip = ExtractCore(skipIndex + 1 & ":" & Ub, ":")
End Function

Public Function Take(skipIndex As Long) As ArrayEx2
    Set Take = ExtractCore(Lb & ":" & skipIndex, ":")
End Function

Public Function Distinct(column As Long) As ArrayEx2
    Dim arr As New ArrayEx1
    Set arr = GetColumn(column)
    Dim i As Long, s As String, sIndexes As String
    For i = arr.Lb To arr.Ub
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
    Dim rearr: ReDim rearr(Lb To Ub)
    Dim i As Long
    For i = Lb To Ub
        rearr(i) = array_(Ub - i + 1)
    Next i
    Dim re As New ArrayEx2
    Set Reverse = ExtractCore(rearr, ":")
End Function

Public Function AllEvaluate(arguments, columnNumbers, formula As String) As Boolean
    AllEvaluate = False
    Dim i As Long, j As Long, s As String, indexes As String, f
    Dim args, cols
    args = Split(arguments, ",")
    cols = Split(columnNumbers, ",")
    If UBound(args) <> UBound(cols) Then Err.Raise 2002, "AllEvaluate", "argument string is unsupported."
    
    For i = Lb To Ub
        s = formula
        For j = LBound(args) To UBound(args)
            s = Replace(s, args(j), array_(i, cols(j)))
        Next j
        If TryEvaluate(s, f) = False Then Exit Function
        If f = False Then Exit Function
    Next i
    AllEvaluate = True
    Exit Function
End Function

Public Function AnyEvaluate(arguments, columnNumbers, formula As String) As Boolean
    AnyEvaluate = True
    Dim i As Long, j As Long, s As String, indexes As String, f
    Dim args, cols
    args = Split(arguments, ",")
    cols = Split(columnNumbers, ",")
    If UBound(args) <> UBound(cols) Then Err.Raise 2002, "AnyEvaluate", "argument string is unsupported."
    
    For i = Lb To Ub
        s = formula
        For j = LBound(args) To UBound(args)
            s = Replace(s, args(j), array_(i, cols(j)))
        Next j
        If TryEvaluate(s, f) = False Then Exit Function
        If f = True Then Exit Function
    Next i
    AnyEvaluate = False
End Function

Public Function WhereEvaluated(arguments, columnNumbers, formula As String) As ArrayEx2
    Dim i As Long, j As Long, s As String, indexes As String, f
    Dim args, cols
    args = Split(arguments, ",")
    cols = Split(columnNumbers, ",")
    If UBound(args) <> UBound(cols) Then Err.Raise 2002, "WhereEvaluated", "argument string is unsupported."
    
    For i = Lb To Ub
        s = formula
        For j = LBound(args) To UBound(args)
            s = Replace(s, args(j), array_(i, cols(j)))
        Next j
        If TryEvaluate(s, f) = False Then Exit Function
        If f Then indexes = indexes & i & ","
    Next i
    indexes = Left(indexes, Len(indexes) - 1)
    Set WhereEvaluated = ExtractCore(indexes, ":")
End Function

Public Function SelectEvaluated(arguments, columnNumbers, formula As String) As ArrayEx1
    Dim i As Long, j As Long, s As String, indexes As String, f
    Dim args, cols
    args = Split(arguments, ",")
    cols = Split(columnNumbers, ",")
    
    Dim rearr()
    ReDim rearr(Lb To Ub)
    For i = Lb To Ub
        s = formula
        For j = LBound(args) To UBound(args)
            s = Replace(s, args(j), array_(i, cols(j)))
        Next j
        If TryEvaluate(s, f, False) = False Then Exit Function
        rearr(i) = f
    Next i

    Dim re As New ArrayEx1
    Set SelectEvaluated = re.Init(rearr)
    Exit Function
End Function

' DebugPrint
Public Function DebugPrint(argWithCommna As String, ColumnNumbersWithConnma As String, formulaWithArg As String) As ArrayEx2
    Dim args: args = Split(Replace(argWithCommna, " ", ""), ",")
    Dim cols: cols = Split(Replace(ColumnNumbersWithConnma, " ", ""), ",")
    Dim i&, j&, s$
    For i = Lb To Ub
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
    Set Transpose = rearr.Init(WorksheetFunction.Transpose(array_))
End Function








' Private Methods --------------------------------------------------------------------------------------------
Private Function ExtractCore(argRows As Variant, argCols As Variant) As ArrayEx2
    Dim rows, cols
    rows = argToArray(argRows, 1)
    cols = argToArray(argCols, 2)

    Dim i, j, ii, jj
    Dim rearr As New ArrayEx2
    Dim re
    ReDim re(1 To UBound(rows) - LBound(rows) + 1, 1 To UBound(cols) - LBound(cols) + 1)
    ii = 1
    For i = LBound(rows) To UBound(rows)
        jj = 1
        For j = LBound(cols) To UBound(cols)
            re(ii, jj) = array_(rows(i), cols(j))
            jj = jj + 1
        Next j
        ii = ii + 1
    Next i
    Set ExtractCore = rearr.Init(re)
End Function

Private Function argToArray(arg As Variant, rowcol As Long) As Variant
    Dim re
    If VarType(arg) = vbLong Or VarType(arg) = vbInteger Then
        argToArray = Array(arg)
        Exit Function
    End If
    If VarType(arg) = vbString Then
        argToArray = DecodeIndexStringToArray(arg, rowcol)
        Exit Function
    End If
    If VarType(arg) > vbArray Then
        argToArray = arg
        Exit Function
    End If
End Function

Private Function DecodeIndexStringToArray(ByVal ToCollone As String, rowcol As Long)
    Dim l As Long, u As Long, l_ As String, u_ As String, t As String, i As Long, j As Long, re, c
    Dim s As String, ss As String, ssarr
    ss = LCase(Replace(ToCollone, " ", ""))
    If ss = ":" Then
        ReDim re(1 To Ub(rowcol) - Lb(rowcol) + 1)
        For i = Lb(rowcol) To Ub(rowcol):             re(i) = i:             Next i
        DecodeIndexStringToArray = re
        Exit Function
    Else
        ssarr = Split(ss, ",")
        Dim collect As New Collection
        For i = LBound(ssarr) To UBound(ssarr)
            s = ssarr(i)
            
            ' Numeric value
            If IsErrorCLng(s) = False Then
                collect.Add CLng(s)
                GoTo Nexti
            End If
            
            ' To, Collone string
            If (InStr(s, "to") > 0 And Len(s) > 2) Or (InStr(s, ":") > 0 And Len(s) > 1) Then
                If InStr(s, "to") > 0 And Len(s) > 2 Then t = "to"
                If InStr(s, ":") > 0 And Len(s) > 2 Then t = ":"
                
                l_ = Left(s, InStr(s, t) - 1)
                u_ = Right(s, InStrRev(s, t) - 1)
                l = IIf(IsNumeric(l_), val(l_), 0)
                u = IIf(IsNumeric(u_), val(u_), 0)
                
                If l = 0 Or u = 0 Then Call Err.Raise(2002, "DecodeIndexStringToArray", "argument string is unsupported.")
                For j = l To u:   collect.Add j:   Next j
                GoTo Nexti
            End If
            
            ' Unsupported string
            Call Err.Raise(2002, "DecodeIndexStringToArray", "argument string is unsupported.")
Nexti:
        Next i

        ReDim re(1 To collect.Count)
        j = 1
        For Each c In collect
            re(j) = c
            j = j + 1
        Next c
        DecodeIndexStringToArray = re
    End If
End Function

Private Function IsErrorCLng(s As String) As Boolean
    On Error GoTo unConvert
    Dim t: t = CLng(s)
    IsErrorCLng = False
    Exit Function
unConvert:
    IsErrorCLng = True
End Function

Private Function TryEvaluate(evaluateString As String, ByRef re As Variant, Optional isBoolean As Boolean = True) As Boolean
    TryEvaluate = False
    re = Evaluate(evaluateString)
    If IsError(re) = True Then GoTo errEvaluate
    If isBoolean And VarType(re) <> vbBoolean Then GoTo errEvaluateFormula
    TryEvaluate = True
    Exit Function

errEvaluate:
    Call Err.Raise(2003, "ArrayEx2, TryEvaluate()", "String in evaluate() is error :" & evaluateString & ".")
errEvaluateFormula:
    Call Err.Raise(2004, "ArrayEx2, TryEvaluate()", "String formula in evaluate() is wrong. should be return boolen. :" & evaluateString & ".")
End Function

Private Function IsArrayRank(inspectedArray As Variant, inspectedRank As Long) As Boolean
    Dim errTest As Long
    Dim i As Long: i = 0
    On Error Resume Next
    Do While Err.Number = 0
        i = i + 1
        errTest = UBound(inspectedArray, i)
    Loop
    On Error GoTo 0
    IsArrayRank = ((i - 1) = inspectedRank)
    Exit Function
End Function


