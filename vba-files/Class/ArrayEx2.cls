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

'Unexpected error, ArrayEx should return instance but null value.

Private array_ As Variant

Private Enum ErrNumAccess
    EC_Public
    EC_Private
End Enum

Private Enum ErrNumType
    EC_NULL
    EC_RANK
    EC_INPUT
    EC_INPUT_SIZE
End Enum

Public Function Init(array2d) As ArrayEx2
    If ArrayRank(array2d) <> 2 Then Call Err.Raise(GetErrNumber(ErrNumAccess.EC_Public, ErrNumType.EC_RANK), "Init", "ArrayEx: need 2d array")
    array_ = array2d
    Set Init = Me
End Function

Public Property Get Value(Optional rows As Variant, Optional cols As Variant)
Attribute Value.VB_UserMemId = 0
    
    If IsMissing(rows) = True And IsMissing(cols) = True Then
        Value = array_
        Exit Property
    End If
    
    If IsMissing(rows) = False And IsMissing(cols) = True Then Value = ExtractCore(rows, ":").ToPrimitive
    If IsMissing(rows) = True And IsMissing(cols) = False Then Value = ExtractCore(":", cols).ToPrimitive
    If IsMissing(rows) = False And IsMissing(cols) = False Then Value = ExtractCore(rows, cols).ToPrimitive
    
End Property

Public Function ToRange(rng As Range)
    rng.Resize(Ub(1), Ub(2)).Value = array_
End Function


Public Function ToPrimitive() As Variant
    If Ub = 1 Then
        If Ub(2) = 1 Then
            ToPrimitive = array_(Ub, Ub(2))
        Else
            ToPrimitive = GetRow(Ub(1)).Value()
        End If
    Else
        If Ub(2) = 1 Then
            ToPrimitive = GetColumn(Ub(2)).Value()
        Else
            ToPrimitive = array_
        End If
    End If
End Function



Public Function SetValue(argRows As Variant, argCols As Variant, argVal As Variant) As ArrayEx2
    Dim rows, cols
    rows = argToArray(argRows, 1)
    cols = argToArray(argCols, 2)

    Dim i, j, ii, jj
    Dim re
    ii = 1
    For i = LBound(rows) To UBound(rows)
        jj = 1
        For j = LBound(cols) To UBound(cols)
            array_(rows(i), cols(j)) = argVal
            jj = jj + 1
        Next j
        ii = ii + 1
    Next i
    Set SetValue = Me
End Function

Public Function SetRow(argRow As Long, argArrayEx1 As ArrayEx1)
    Dim i
    If Lb(2) <> argArrayEx1.Lb And Ub(2) <> argArrayEx1.Ub Then _
        Call Err.Raise(GetErrNumber(ErrNumAccess.EC_Public, ErrNumType.EC_INPUT_SIZE), "SetRow", "ArrayEx: need same length")
    
    Dim re:  re = argArrayEx1.Value
    For i = Lb(2) To Ub(2)
        array_(argRow, i) = re(i)
    Next i
End Function

Public Function SetColumn(argCol As Long, argArrayEx1 As ArrayEx1)
    Dim i
    If Lb(1) <> argArrayEx1.Lb And Ub(1) <> argArrayEx1.Ub Then _
        Call Err.Raise(GetErrNumber(ErrNumAccess.EC_Public, ErrNumType.EC_INPUT_SIZE), "SetColumn", "ArrayEx: need same length")
    
    Dim re:  re = argArrayEx1.Value
    For i = Lb(1) To Ub(1)
        array_(i, argCol) = re(i)
    Next i
End Function


Public Function Transpose() As ArrayEx2
    Dim rearr As New ArrayEx2
    Set Transpose = rearr.Init(WorksheetFunction.Transpose(array_))
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

Public Function Extract(rows As Variant, cols As Variant) As ArrayEx2
    Set Extract = ExtractCore(rows, cols)
End Function

Private Function argToArray(arg As Variant, rowcol As Long) As Variant
Dim re
If VarType(arg) = vbLong Or VarType(arg) = vbInteger Then
    argToArray = Array(arg)
    Exit Function
End If
If VarType(arg) = vbString Then
    argToArray = DecodeToCollone2Array(arg, rowcol)
    Exit Function
End If
If VarType(arg) > vbArray Then
    argToArray = arg
    Exit Function
End If
End Function

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

Public Function CreateFromCollection(col) As ArrayEx2
    Dim rearr As New ArrayEx2
    Dim row As New ArrayEx1
    Dim re
    Dim c
    For Each c In col
    If ArrayRank(c) <> 1 Then Exit Function
    Set rearr = rearr.AddRow(row.Create(c))
    Next c
    Set CreateFromCollection = rearr
End Function

Public Property Get Ub(Optional direction As Long = 1) As Long
    Ub = UBound(array_, direction)
End Property

Public Property Get Lb(Optional direction As Long = 1) As Long
    Lb = LBound(array_, direction)
End Property



Public Function DebugPrintEvaluated(argWithCommna As String, ColumnNumbersWithConnma As String, formulaWithArg As String) As ArrayEx2
    Dim args: args = Split(Replace(argWithCommna, " ", ""), ",")
    Dim cols: cols = Split(Replace(ColumnNumbersWithConnma, " ", ""), ",")
    Dim i&, j&, s$
    For i = Lb To Ub
        s = formulaWithArg
        For j = LBound(args) To UBound(args)
            s = Replace(s, args(j), array_(i, val(cols(j))))
        Next j
       Debug.Print Evaluate(s)
    Next i
    Set DebugPrintEvaluated = Me
End Function

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
    Dim i&, j&, s$
    For i = Lb To Ub
        For j = Lb(2) To Ub(2)
            s = s & array_(i, j) & ","
            Next j
            s = s & vbCrLf
    Next i
       Debug.Print s
    Set DebugPrintAll = Me
End Function


Function AddRow(row) As ArrayEx2
    Dim i, j
    Dim re
  
     If IsEmpty(array_) = False Then
         ReDim re(Lb To Ub + 1, Lb(2) To Ub(2))
          For i = Lb To Ub
          For j = Lb(2) To Ub(2)
                re(i, j) = array_(i, j)
            Next j
        Next i
        For j = Lb(2) To Ub(2)
            re(Ub + 1, j) = row(j)
        Next j
    Else
        ReDim re(1 To 1, LBound(row) To UBound(row))
        For j = LBound(re, 2) To UBound(re, 2)
            re(1, j) = row(j)
        Next j
    End If
        Dim rearr As New ArrayEx2
        Set AddRow = rearr.Init(re)
End Function






Private Function DecodeToCollone2Array(ByVal ToCollone As String, rowcol As Long)
Dim l, u, i, l_, u_, t
Dim re, c
Dim s As String, ss As String, ssarr
    ss = LCase(Replace(ToCollone, " ", ""))
    If ss = ":" Then
        ReDim re(1 To Ub(rowcol) - Lb(rowcol) + 1)
        For i = Lb(rowcol) To Ub(rowcol):             re(i) = i:             Next i
        DecodeToCollone2Array = re
        Exit Function
    Else
        ssarr = Split(ss, ",")
        Dim collect As New Collection
        Dim j
        For i = LBound(ssarr) To UBound(ssarr)
            s = ssarr(i)
            If IsErrorCLng(s) = False Then
                collect.Add CLng(s)
                GoTo Nexti
            End If
            
            If (InStr(s, "to") > 0 And Len(s) > 2) Or (InStr(s, ":") > 0 And Len(s) > 1) Then
                If InStr(s, "to") > 0 And Len(s) > 2 Then t = "to"
                If InStr(s, ":") > 0 And Len(s) > 2 Then t = ":"
                
                l_ = Left(s, InStr(s, t) - 1)
                u_ = Right(s, InStrRev(s, t) - 1)
                l = IIf(IsNumeric(l_), val(l_), 0)
                u = IIf(IsNumeric(u_), val(u_), 0)
                
                If l = 0 Or u = 0 Then
                    Call Err.Raise(GetErrNumber(ErrNumAccess.EC_Private, ErrNumType.EC_INPUT), _
                    "DecodeToCollone2Array", "argument string is unsported.")
                End If
                For j = l To u:                  collect.Add j:             Next j
                GoTo Nexti
            End If
            
            Call Err.Raise(GetErrNumber(ErrNumAccess.EC_Private, ErrNumType.EC_INPUT), _
                    "DecodeToCollone2Array", "argument string is unsported.")
Nexti:
        Next i
        
        ReDim re(1 To collect.Count)
        j = 1
        For Each c In collect
            re(j) = c
            j = j + 1
        Next c
        DecodeToCollone2Array = re
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


Private Function GetErrNumber(Optional accessLv As Long = 0, Optional errType As Long = 0) As Integer
    GetErrNumber = 2 * 1000 + accessLv * 100 + errType
End Function

Private Function errMsg(Optional errType As Long = 0) As String
    errMsg = ": need 2d array"
End Function


' [Get] îzóÒÇÃéüå≥êî
Private Function ArrayRank(Optional ary As Variant = Empty) As Integer
    Dim arr As Variant: arr = IIf(IsEmpty(ary), array_, ary)
    On Error Resume Next
    Dim errTest As Long, i As Long
    Do While Err.Number = 0
        i = i + 1
        errTest = UBound(arr, i)
    Loop
    On Error GoTo 0
    ArrayRank = i - 1
End Function
