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
Public Property Get Value(Optional indexes As Variant)
Attribute Value.VB_UserMemId = 0
    If IsMissing(indexes) = True Then
        Value = array_
    Else
        Value = ExtractCore(indexes).ToPrimitive
    End If
End Property

' [Hack] need to more good idea.
' for argument.  need for return value as ArrayEx1
Public Property Get Mine()
    Set Mine = Me
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
    If IsNull(array1d) Then array_ = Null: Exit Function:
    If IsArrayRank(array1d, 1) = False Then Call Err.Raise(1101, "ArrayEx1 Init", "ArrayEx: Need 1d array")
    array_ = array1d
    Set Init = Me
End Function

Public Function Create(Optional array1d As Variant) As ArrayEx1
    If IsMissing(array1d) = True And IsEmpty(array_) Then Call Err.Raise(1105, "ArrayEx1 Create", "ArrayEx: Not initialized")
    Dim rearr As New ArrayEx1
    If IsMissing(array1d) Then
        Set Create = rearr.Init(array_)
    Else
        Set Create = rearr.Init(array1d)
    End If
End Function

Public Function Equals(array1d) As Boolean
    Equals = False
    If Lb <> LBound(array1d) Then Exit Function
    If Ub <> UBound(array1d) Then Exit Function
    Dim i As Long
    For i = Lb To Ub
        If array_(i) <> array1d(i) Then Exit Function
    Next i
    Equals = True
End Function

Public Function ToPrimitive() As Variant
    If Ub = 1 Then ToPrimitive = array_(1)
    If Ub <> 1 Then ToPrimitive = array_
End Function

' Set
Public Function SetElement(index As Long, argVal) As ArrayEx1
    array_(index) = argVal
    Set SetElement = Me
End Function

' Get
Public Function GetElements(indexes) As ArrayEx1
    Set GetElements = ExtractCore(indexes)
End Function

Public Function GetElement(index As Long) As ArrayEx0
    Dim rearr As New ArrayEx0
    Set GetElement = rearr.Init(array_(index))
End Function

' Add
Function AddElement(element As ArrayEx0) As ArrayEx1
    Dim i, j
    Dim re
    If IsEmpty(array_) Then
        ReDim re(1)
        re(1) = element.Value
    Else
        ReDim re(Lb To Ub + 1)
        For i = Lb To Ub
            re(i) = array_(i)
        Next i
        re(Ub + 1) = element.Value
    End If
    Dim rearr As New ArrayEx1
    Set AddElement = rearr.Init(re)
End Function

' Convert
Public Function ToString() As String
    Dim s As String, i As Long, j As Long
    s = "["
    For i = Lb To Ub
        s = s & CStr(array_(i)) & ","
    Next i
    s = Left(s, Len(s) - 1) & "]"
    ToString = s
End Function

Public Function ToRange(rng As Range)
    rng.Resize(1, Ub).Value = array_
End Function

Public Function ToCollection() As Collection
    Dim coll As New Collection
    Dim arr As New ArrayEx0
    
    Dim i As Long
    For i = Lb To Ub
        coll.Add arr.Create(array_(i))
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
    Set Skip = ExtractCore(skipIndex + 1 & ":" & Ub)
End Function

Public Function Take(skipIndex As Long) As ArrayEx1
    Set Take = ExtractCore(Lb & ":" & skipIndex)
End Function

Public Function Distinct() As ArrayEx1
    Dim re As New ArrayEx1
    Dim var As Variant
    Dim i As Long, s As String, sIndexes As String
    For i = Lb To Ub
        If InStr(s, array_(i)) = 0 Then
            s = s & array_(i) & ","
            sIndexes = sIndexes & i & ","
        End If
    Next i
    If sIndexes = "" Then Exit Function
    sIndexes = Left(sIndexes, Len(sIndexes) - 1)
    Set Distinct = ExtractCore(sIndexes)
End Function

Public Function OrderBy() As ArrayEx1
    Dim re As New ArrayEx1
    Set OrderBy = re.Init(WorksheetFunction.Sort(array_))
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
    Dim i As Long, f
    For i = Lb To Ub
        If TryEvaluate(Replace(formula, argument, array_(i)), f) = False Then Exit Function
        If f = False Then Exit Function
    Next i
    AllEvaluate = True
    Exit Function
End Function

Public Function AnyEvaluate(argument As String, formula As String) As Boolean
    AnyEvaluate = True
    Dim i As Long, f
    For i = Lb To Ub
        If TryEvaluate(Replace(formula, argument, array_(i)), f) = False Then Exit Function
        If f = True Then Exit Function
    Next i
    AnyEvaluate = False
End Function

Public Function WhereEvaluated(argument As String, formula As String) As ArrayEx1
    Dim i As Long, s As String, indexes As String, f
    For i = Lb To Ub
        s = Replace(formula, argument, array_(i))
        If TryEvaluate(s, f) = False Then Exit Function
        If f Then indexes = indexes & i & ","
    Next i
    indexes = Left(indexes, Len(indexes) - 1)
    Set WhereEvaluated = ExtractCore(indexes)
End Function

Public Function SelectEvaluated(argument As String, formula As String) As ArrayEx1
    Dim i As Long, rearr(), s As String
    ReDim rearr(Lb To Ub)
    For i = Lb To Ub
        s = Replace(formula, argument, array_(i))
        rearr(i) = Evaluate(s)
        If IsError(rearr(i)) = True Then GoTo errEvaluate
    Next i
    Dim rearr0 As New ArrayEx1
    Set SelectEvaluated = rearr0.Init(rearr)
    Exit Function
errEvaluate:
    Call Err.Raise(1103, "ArrayEx1, SelectEvaluated()", "String in evaluate() is error :" & s & ".")
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


' Private Methods ------------------------------------------------------------------
Private Function ExtractCore(argIndexs As Variant) As ArrayEx1
    Dim rows
    rows = argToArray(argIndexs)
    Dim i As Long, ii As Long
    Dim re
    ReDim re(1 To UBound(rows) - LBound(rows) + 1)
    ii = 1
    For i = LBound(rows) To UBound(rows)
        re(ii) = array_(rows(i))
        ii = ii + 1
    Next i
    Dim rearr As New ArrayEx1
    Set ExtractCore = rearr.Init(re)
End Function

Private Function argToArray(arg As Variant) As Variant
    Dim re
    If VarType(arg) = vbLong Or VarType(arg) = vbInteger Then
        argToArray = Array(arg)
        Exit Function
    End If
    If VarType(arg) = vbString Then
        argToArray = DecodeIndexStringToArray(arg)
        Exit Function
    End If
    If VarType(arg) > vbArray Then
        argToArray = arg
        Exit Function
    End If
End Function

Private Function DecodeIndexStringToArray(ByVal ToCollone As String)
    Dim l As Long, u As Long, l_ As String, u_ As String, t As String, i As Long, j As Long, re, c
    Dim s As String, ss As String, ssarr
    ss = LCase(Replace(ToCollone, " ", ""))
    If ss = ":" Then
        ReDim re(1 To Ub() - Lb() + 1)
        For i = Lb() To Ub():             re(i) = i:             Next i
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

Private Function TryEvaluate(evaluateString As String, ByRef re As Variant) As Boolean
    TryEvaluate = False
    re = Evaluate(evaluateString)
    If IsError(re) = True Then GoTo errEvaluate
    If VarType(re) <> vbBoolean Then GoTo errEvaluateFormula
    TryEvaluate = True
    Exit Function

errEvaluate:
    Call Err.Raise(1003, "ArrayEx1, TryEvaluate()", "String in evaluate() is error :" & evaluateString & ".")
errEvaluateFormula:
    Call Err.Raise(1004, "ArrayEx1, TryEvaluate()", "String formula in evaluate() is wrong. should be return boolen. :" & evaluateString & ".")
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

