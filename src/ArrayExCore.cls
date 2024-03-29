VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayExCore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'todo Done そもそもcoreにフィールドとして入れると、ExtractCoreとかで更新されるときに、都度更新しないといけないので、困る．．．
'todo Done 添え字は０でなくてもよくする。が、１出ないとダメなわけでもない。　可変
'todo Done 添え字をシフトするメソッドの追加
'todo Done Value は素直に配列を返す
'todo Done VSTACK, HSTACK の追加
'todo その他、新関数をArrayExCoreに
'todo ExtractCoreがArrayExCoreに入るか・・・？
'todo 要素をCollectionに
'todo ParamArrayに出来るとこを,ParamArrayに
'hack 左結合 LeftJoin
'todo　Done ”:”, "," は定数化
'todo　debug.print,  toStringのカッコは [ -> { ?
'hack 全ての組合せを求める
' todo error message
'todo Done  Hstack

Private Const COLLON As String = ":"
Private Const CONNMA As String = ","
Private Const INDEX_TO As String = "to"


Public Function Validate(arr)
    If IsEmpty(arr) Then Err.Raise 9100, "ArrayExCore.Validate", "argument is invalid. (Empty)"
    If IsError(arr) Then Err.Raise 9100, "ArrayExCore.Validate", "argument is invalid. (Error)"
    If IsNull(arr) Then Err.Raise 9100, "ArrayExCore.Validate", "argument is invalid.(Null)"
    If VarType(arr) Mod vbArray = vbObject Then Err.Raise 9100, "ArrayExCore.Validate", "argument is invalid.(Object)"
End Function

Public Function Rank(arr)
    Dim i As Long, errTest As Long
    On Error GoTo outLoop
    i = 0
    Do While Err.Number = 0
        i = i + 1
        errTest = UBound(arr, i)
        If i > 3 Then GoTo outLoop
    Loop
outLoop:
    On Error GoTo 0
    Err.Clear
    Rank = (i - 1)
End Function

Public Function ConvertToDimensionN(arr, dimension As Long)
    Dim i As Long, j As Long, rnk As Long, lb1 As Long, ub1 As Long, lb2 As Long, ub2 As Long, cnt1 As Long, cnt2 As Long
    Dim rearr, key As Long
    
    Validate arr
    GetArrayParameters arr, rnk, lb1, ub1, cnt1, lb2, ub2, cnt2
    key = 10 * rnk + dimension
    Select Case key
        Case 0:     rearr = arr
        Case 1:     ReDim rearr(0): rearr(0) = arr
        Case 2:     ReDim rearr(0, 0): rearr(0, 0) = arr
        Case 10:    rearr = arr(lb1)
        Case 11:    rearr = arr
        Case 12:   ReDim rearr(0, lb1 To ub1)
                        For i = lb1 To ub1: rearr(0, i) = arr(i): Next i
        Case 20:    rearr = arr(lb1, lb1)
        Case 21:   ReDim rearr(lb2 To ub2)
                        For i = lb2 To ub2: rearr(i) = arr(lb1, i): Next i
        Case 22:    rearr = arr
        Case Else
            Call Err.Raise(9101, "ArrayExCore.ConvertToDimensionN", "ArrayEx: not supported dimension")
    End Select
    ConvertToDimensionN = rearr
End Function

Public Function RedimPreserve2(arr, lbound1 As Long, ubound1 As Long, lbound2 As Long, ubound2 As Long, Optional blnk As Variant = "") As Variant()
    Dim i As Long, j As Long, rnk As Long, lb1 As Long, ub1 As Long, lb2 As Long, ub2 As Long, cnt1 As Long, cnt2 As Long
    Dim rearr As Variant, cnti, cntj
    
    ReDim rearr(lbound1 To ubound1, lbound2 To ubound2)
    SetValueToAllElement rearr, blnk
    GetArrayParameters arr, rnk, lb1, ub1, cnt1, lb2, ub2, cnt2
    For cnti = 0 To ubound1 - lbound1
        If ub1 < lb1 + cnti Then GoTo outLoop
        For cntj = 0 To ubound2 - lbound2
            If ub2 < lb2 + cntj Then GoTo nextj
            rearr(lbound1 + cnti, lbound2 + cntj) = arr(lb1 + cnti, lb2 + cntj)
        Next cntj
nextj:
    Next cnti
outLoop:
        
    RedimPreserve2 = rearr
End Function

Public Function TryEvaluate(evaluateString As String, ByRef re As Variant) As Boolean
    TryEvaluate = False: re = Null
On Error GoTo errEvaluate
    re = Evaluate(evaluateString)
    If IsError(re) = True Then GoTo errEvaluate
    TryEvaluate = True
errEvaluate:
End Function

Public Function ExtractCore(arr, Indexes1 As Variant, Optional Indexes2 As Variant) As Variant
    Dim i As Long, j As Long, rnk As Long, lb1 As Long, ub1 As Long, lb2 As Long, ub2 As Long, cnt1 As Long, cnt2 As Long
    Dim index1, index2, ii, jj, re
    
    GetArrayParameters arr, rnk, lb1, ub1, cnt1, lb2, ub2, cnt2
    ii = lb1: jj = lb2
    Select Case rnk
    Case 0
        re = arr
    Case 1
        index1 = IndexToArray(Indexes1, lb1, ub1)
        ReDim re(lb1 To lb1 + UBound(index1) - LBound(index1))
        For i = LBound(index1) To UBound(index1): re(ii) = arr(index1(i)): ii = ii + 1: Next i
    Case 2
        index1 = IndexToArray(Indexes1, lb1, ub1)
        index2 = IndexToArray(Indexes2, lb2, ub2)
        ReDim re(lb1 To lb1 + UBound(index1) - LBound(index1), lb2 To lb2 + UBound(index2) - LBound(index2))
        For i = LBound(index1) To UBound(index1)
            jj = lb2
            For j = LBound(index2) To UBound(index2)
                re(ii, jj) = arr(index1(i), index2(j))
                jj = jj + 1
            Next j
            ii = ii + 1
        Next i
    Case Else
    End Select
    ExtractCore = re
End Function

Public Function IndexToArray(index As Variant, low As Long, up As Long) As Variant
    Dim re
    If VarType(index) = vbLong Or VarType(index) = vbInteger Then
        IndexToArray = Array(index)
        Exit Function
    End If
    If VarType(index) = vbString Then
        IndexToArray = DecodeIndexStringToArray(index, low, up)
        Exit Function
    End If
    If VarType(index) > vbArray Then
        IndexToArray = index
        Exit Function
    End If
    If VarType(index) = vbObject Then
        If TypeName(index) = "Collection" Then: Dim coll As Collection: Set coll = index: _
            IndexToArray = GetFromCollection(coll)
    End If
End Function

Public Function GetFromSpillingRange(rng As Range)
    If rng.HasSpill = False Then GetFromSpillingRange = ConvertToDimensionN(rng.Value, 2): Exit Function
    GetFromSpillingRange = rng.SpillingToRange.Value
End Function

Public Function ShiftIndex(arr, addIndex1 As Long, Optional addIndex2 As Long = 0)
    Dim i As Long, j As Long, rnk As Long, lb1 As Long, ub1 As Long, lb2 As Long, ub2 As Long, cnt1 As Long, cnt2 As Long
    Dim rearr
    
    Validate arr
    GetArrayParameters arr, rnk, lb1, ub1, cnt1, lb2, ub2, cnt2
    If rnk = 0 Then Err.Raise "ArrayExCore.ShiftIndex", "Need dimension more 0"
    If rnk > 0 Then If lb1 + addIndex1 < 0 Then Err.Raise 9001, "ArrayExCore.ShiftIndex", "Need lbound1 index 0 or more"
    If rnk > 1 Then If lb2 + addIndex2 < 0 Then Err.Raise 9001, "ArrayExCore.ShiftIndex", "Need lbound2 index 0 or more"

    If rnk = 1 Then
        ReDim rearr(lb1 + addIndex1 To ub1 + addIndex1)
        For i = lb1 + addIndex1 To ub1 + addIndex1:  rearr(i) = arr(i - lb1): Next i
    End If
    If rnk = 2 Then
        rearr = RedimPreserve2(arr, lb1 + addIndex1, ub1 + addIndex1, lb2 + addIndex2, ub2 + addIndex2)
    End If
    
    ShiftIndex = rearr
End Function


' WorksheetFucntions extended -----------------------

Public Function TEXTBEFORE(ByVal text As String, delimiter, Optional instance_num As Long = 1, Optional match_mode As Long = 0, Optional match_end As Long = 0, Optional if_not_found As Variant)
    Dim str, strIndex As Long, txt As String, col
    Dim delimiters As Collection: Set delimiters = New Collection
    
    If match_mode = 1 Then text = LCase(text): delimiter = LCase(delimiter)
    If VarType(delimiter) = vbString Then
        delimiters.Add delimiter
    Else
        For Each str In delimiter:                 delimiters.Add str:        Next
    End If
    txt = text
    Set col = GetSplitStringCollection(txt, delimiters, False, match_mode, instance_num)
    
    If IsMissing(if_not_found) Then if_not_found = CVErr(xlErrNA)
    If col.Count = 1 Then TEXTBEFORE = if_not_found: Exit Function
    If col.Count < instance_num + 1 Then TEXTBEFORE = CVErr(xlErrNA): Exit Function
    
    TEXTBEFORE = col(IIf(instance_num > 0, instance_num, instance_num * -1))
End Function
Public Function TEXTAFTER(ByVal text As String, delimiter, Optional instance_num As Long = 1, Optional match_mode As Long = 0, Optional match_end As Long = 0, Optional if_not_found As Variant)
    Dim str, strIndex As Long, txt As String, col
    Dim delimiters As Collection: Set delimiters = New Collection
    
    If match_mode = 1 Then text = LCase(text): delimiter = LCase(delimiter)
    If VarType(delimiter) = vbString Then
        delimiters.Add delimiter
    Else
        For Each str In delimiter:                 delimiters.Add str:        Next
    End If
    txt = text
    Set col = GetSplitStringCollection(txt, delimiters, False, match_mode, instance_num)
    
    If IsMissing(if_not_found) Then if_not_found = CVErr(xlErrNA)
    If col.Count = 1 Then TEXTAFTER = if_not_found: Exit Function
    If col.Count < instance_num + 1 Then TEXTAFTER = CVErr(xlErrNA): Exit Function
    
    TEXTAFTER = col(IIf(instance_num > 0, instance_num + 1, instance_num * -1))
End Function
Public Function TEXTSPLIT(ByVal text As String, col_delimiter As Variant, Optional row_delimiter, Optional ignore_empty As Boolean = False, Optional match_mode As Long = 0, Optional pad_with As Variant)
    Dim col As Collection:    Set col = New Collection
    Dim rw As Collection:    Set rw = New Collection
    Dim cllec As Collection: Set cllec = New Collection
    Dim columnDelimiters As Collection: Set columnDelimiters = New Collection
    Dim rowDelimiters As Collection: Set rowDelimiters = New Collection
    Dim str, txt As String, co, maxcount As Long, strcol As New Collection
    
    If IsMissing(pad_with) Then pad_with = CVErr(xlErrNA)
    If VarType(col_delimiter) = vbString Then
        columnDelimiters.Add col_delimiter
    Else
        For Each str In col_delimiter:                 columnDelimiters.Add str:        Next
    End If
    If VarType(row_delimiter) = vbString Then
        rowDelimiters.Add row_delimiter
    Else
        For Each str In row_delimiter:                  rowDelimiters.Add str:        Next
    End If
    txt = text
    Set col = GetSplitStringCollection(txt, rowDelimiters, ignore_empty, match_mode)
    For Each co In col
        strcol.Add GetSplitStringCollection(CStr(co), columnDelimiters, ignore_empty, match_mode)
    Next
    maxcount = GetMaxCount(strcol)
    For Each co In strcol
        cllec.Add RedimCollectionToArray2D(co, 1, 1, 1, maxcount, pad_with)
    Next

    TEXTSPLIT = StackArray(cllec, False)
End Function
Public Function HSTACK(array1, ParamArray paramarrs())
    Dim paramarr
    Dim arrays As Collection: Set arrays = New Collection

    arrays.Add ConvertToDimensionN(array1, 2)
    For Each paramarr In paramarrs
        arrays.Add ConvertToDimensionN(paramarr, 2)
    Next
    
    HSTACK = StackArray(arrays, True)
End Function
Public Function VSTACK(array1, ParamArray paramarrs())
    Dim paramarr
    Dim arrays As Collection: Set arrays = New Collection

    arrays.Add ConvertToDimensionN(array1, 2)
    For Each paramarr In paramarrs
        arrays.Add ConvertToDimensionN(paramarr, 2)
    Next
    
    VSTACK = StackArray(arrays, False)
End Function
Public Function TOROW(arr, Optional ignore As Long = 0, Optional scan_by_column As Boolean = False)
    Dim i As Long, j As Long, rnk As Long, lb1 As Long, ub1 As Long, lb2 As Long, ub2 As Long, cnt1 As Long, cnt2 As Long
    Dim rearr, ele As Variant, coll As Collection
    Set coll = New Collection
    
    GetArrayParameters arr, rnk, lb1, ub1, cnt1, lb2, ub2, cnt2
    arr = ConvertToDimensionN(arr, 2)
    For Each ele In GetElementsCollection(arr, Not (scan_by_column))
        If ignore = 1 And IsEmpty(ele) Then GoTo nextStep
        If ignore = 2 And IsError(ele) Then GoTo nextStep
        If ignore = 3 And (IsEmpty(ele) Or IsError(ele)) Then GoTo nextStep
        coll.Add ele
nextStep:
    Next
    
    TOROW = RedimCollectionToArray2D(coll, lb1, lb1, lb2, lb2 + coll.Count - 1)
End Function
Public Function TOCOL(arr, Optional ignore As Long = 0, Optional scan_by_column As Boolean = False)
    Dim i As Long, j As Long, rnk As Long, lb1 As Long, ub1 As Long, lb2 As Long, ub2 As Long, cnt1 As Long, cnt2 As Long
    Dim rearr, ele As Variant, coll As Collection
    
    Set coll = New Collection
    GetArrayParameters arr, rnk, lb1, ub1, cnt1, lb2, ub2, cnt2
    arr = ConvertToDimensionN(arr, 2)
    For Each ele In GetElementsCollection(arr, Not (scan_by_column))
        If ignore = 1 And IsEmpty(ele) Then GoTo nextStep
        If ignore = 2 And IsError(ele) Then GoTo nextStep
        If ignore = 3 And (IsEmpty(ele) Or IsError(ele)) Then GoTo nextStep
        coll.Add ele
nextStep:
    Next
       
    TOCOL = RedimCollectionToArray2D(coll, lb1, lb1 + coll.Count - 1, lb2, lb2)
End Function
Public Function WRAPROW(ByVal vector, wrap_count As Long, Optional pad_with As Variant)
    Dim i As Long, j As Long, rnk As Long, lb1 As Long, ub1 As Long, lb2 As Long, ub2 As Long, cnt1 As Long, cnt2 As Long
    
    GetArrayParameters vector, rnk, lb1, ub1, cnt1, lb2, ub2, cnt2
    vector = ConvertToDimensionN(vector, 2)
    If lb1 <> ub1 And lb2 <> ub2 Then WRAPROW = CVErr(xlErrValue): Exit Function
    If wrap_count < 1 Then WRAPROW = CVErr(xlErrNum): Exit Function
    If IsMissing(pad_with) Then pad_with = CVErr(xlErrNA)
    lb2 = lb1
    
    WRAPROW = RedimCollectionToArray2D(vector, lb1, lb1 + (Fix(IIf(cnt1 > cnt2, cnt1, cnt2) / wrap_count) + 1) - 1, lb2, lb2 + wrap_count - 1, pad_with)
End Function
Public Function WRAPCOLS(ByVal vector, wrap_count As Long, Optional pad_with As Variant)
    Dim i As Long, j As Long, rnk As Long, lb1 As Long, ub1 As Long, lb2 As Long, ub2 As Long, cnt1 As Long, cnt2 As Long
    
    GetArrayParameters vector, rnk, lb1, ub1, cnt1, lb2, ub2, cnt2
    vector = ConvertToDimensionN(vector, 2)
    If lb1 <> ub1 And lb2 <> ub2 Then WRAPCOLS = CVErr(xlErrValue): Exit Function
    If wrap_count < 1 Then WRAPCOLS = CVErr(xlErrNum): Exit Function
    If IsMissing(pad_with) Then pad_with = CVErr(xlErrNA)
    lb2 = lb1
    
    WRAPCOLS = RedimCollectionToArray2D(vector, lb1, lb1 + wrap_count - 1, lb2, lb2 + (Fix(IIf(cnt1 > cnt2, cnt1, cnt2) / wrap_count) + 1) - 1, pad_with, False)
End Function
Public Function EXPAND(arr, Optional rows As Variant, Optional columns As Variant, Optional pad_with As Variant)
    Dim i As Long, j As Long, rnk As Long, lb1 As Long, ub1 As Long, lb2 As Long, ub2 As Long, cnt1 As Long, cnt2 As Long
    
    GetArrayParameters arr, rnk, lb1, ub1, cnt1, lb2, ub2, cnt2
    arr = ConvertToDimensionN(arr, 2)
    If IsMissing(rows) Then rows = ub1
    If IsMissing(columns) Then columns = ub2
    If rows < cnt1 Or columns < cnt2 Then Err.Raise 9001, "EXPAND", "Argument is smaller than array size"
    If IsMissing(pad_with) Then pad_with = CVErr(xlErrNA)
            
    EXPAND = RedimPreserve2(arr, lb1, lb1 + rows - 1, lb2, lb2 + columns - 1, pad_with)
End Function
Public Function TAKE(arr, Optional rows As Variant, Optional columns As Variant) As Variant
    Dim i As Long, j As Long, rnk As Long, lb1 As Long, ub1 As Long, lb2 As Long, ub2 As Long, cnt1 As Long, cnt2 As Long
    Dim c1 As Collection: Set c1 = New Collection
    Dim c2 As Collection: Set c2 = New Collection
    
    GetArrayParameters arr, rnk, lb1, ub1, cnt1, lb2, ub2, cnt2
    arr = ConvertToDimensionN(arr, 2)
    If IsMissing(rows) = False Then If rows = 0 Then TAKE = CVErr(xlErrCalc): Exit Function
    If IsMissing(columns) = False Then If columns = 0 Then TAKE = CVErr(xlErrCalc): Exit Function
    If IsMissing(rows) Then rows = ub1
    If IsMissing(columns) Then columns = ub2
    
    For i = lb1 To ub1
        If rows >= 0 And i - rows <= 0 Then c1.Add i
        If rows < 0 And i - (ub1 + rows) > 0 Then c1.Add i
    Next i
    For j = lb2 To ub2
        If columns >= 0 And j - columns <= 0 Then c2.Add j
        If columns < 0 And j - (ub2 + columns) > 0 Then c2.Add j
    Next j
    
    TAKE = ExtractCore(arr, c1, c2)
End Function
Public Function DROP(arr, Optional rows As Variant, Optional columns As Variant) As Variant
    Dim i As Long, j As Long, rnk As Long, lb1 As Long, ub1 As Long, lb2 As Long, ub2 As Long, cnt1 As Long, cnt2 As Long
    Dim c1 As Collection: Set c1 = New Collection
    Dim c2 As Collection: Set c2 = New Collection
    
    GetArrayParameters arr, rnk, lb1, ub1, cnt1, lb2, ub2, cnt2
    arr = ConvertToDimensionN(arr, 2)
    If IsMissing(rows) = False Then If rows = 0 Then DROP = CVErr(xlErrCalc): Exit Function
    If IsMissing(columns) = False Then If columns = 0 Then DROP = CVErr(xlErrCalc): Exit Function
    If IsMissing(rows) Then rows = lb1
    If IsMissing(columns) Then columns = lb2
    
    For i = lb1 To ub1
        If rows >= 0 And i - rows > 0 Then c1.Add i
        If rows < 0 And i - (ub1 + rows) > 0 Then c1.Add i
    Next i
    For j = lb2 To ub2
        If columns >= 0 And j - columns > 0 Then c2.Add j
        If columns < 0 And j - (ub2 + columns) > 0 Then c2.Add j
    Next j
    
    DROP = ExtractCore(arr, c1, c2)
End Function
Public Function CHOOSEROWS(arr, ParamArray paramarr()) As Variant()
    Dim i As Long, j As Long, rnk As Long, lb1 As Long, ub1 As Long, lb2 As Long, ub2 As Long, cnt1 As Long, cnt2 As Long
    Dim indexes As Collection: Set indexes = New Collection
    Dim indx As Long
    
    GetArrayParameters arr, rnk, lb1, ub1, cnt1, lb2, ub2, cnt2
    arr = ConvertToDimensionN(arr, 2)
    For i = LBound(paramarr, 1) To UBound(paramarr, 1)
        If paramarr(i) = 0 Or paramarr(i) > cnt1 Then CHOOSEROWS = CVErr(xlErrValue): Exit Function
        If paramarr(i) > 0 Then indx = lb1 + paramarr(i) - 1 Else indx = ub1 + paramarr(i)
        indexes.Add indx
    Next i
    
    CHOOSEROWS = ExtractCore(arr, indexes, ":")
End Function
Public Function CHOOSECOLS(arr, ParamArray paramarr()) As Variant()
    Dim i As Long, j As Long, rnk As Long, lb1 As Long, ub1 As Long, lb2 As Long, ub2 As Long, cnt1 As Long, cnt2 As Long
    Dim indexes As Collection: Set indexes = New Collection
    Dim indx As Long
    
    GetArrayParameters arr, rnk, lb1, ub1, cnt1, lb2, ub2, cnt2
    arr = ConvertToDimensionN(arr, 2)
    For i = LBound(paramarr, 1) To UBound(paramarr, 1)
        If paramarr(i) = 0 Or paramarr(i) > cnt2 Then CHOOSECOLS = CVErr(xlErrValue): Exit Function
        If paramarr(i) > 0 Then indx = lb2 + paramarr(i) - 1 Else indx = ub2 + paramarr(i)
        indexes.Add indx
    Next i
    
    CHOOSECOLS = ExtractCore(arr, ":", indexes)
End Function



' Private --------------------------------------------------------------------------------------------------
Private Function DecodeIndexStringToArray(ByVal indexString As String, low As Long, up As Long)
    Dim l As Long, u As Long, l_ As String, u_ As String, t As String, i As Long, j As Long, re, c
    Dim s As String, ss As String, ssarr
    
    ss = LCase(Replace(indexString, " ", ""))
    If Right(ss, 1) = CONNMA Then ss = Left(ss, Len(ss) - 1)
    
    If ss = COLLON Then
        ReDim re(low To up)
        For i = low To up
            re(i) = i
        Next i
        DecodeIndexStringToArray = re
        Exit Function
    Else
        ssarr = Split(ss, CONNMA)
        Dim collect As New Collection
        For i = LBound(ssarr) To UBound(ssarr)
            s = ssarr(i)
            
            ' Numeric value
            If IsErrorCLng(s) = False Then
                collect.Add CLng(s)
                GoTo nexti
            End If
            
            ' To, Collone string
            If (InStr(s, INDEX_TO) > 0 And Len(s) > 2) Or (InStr(s, COLLON) > 0 And Len(s) > 1) Then
                If InStr(s, INDEX_TO) > 0 And Len(s) > 2 Then t = INDEX_TO
                If InStr(s, COLLON) > 0 And Len(s) > 2 Then t = COLLON
                
                l_ = Left(s, InStr(s, t) - 1)
                u_ = Right(s, InStrRev(s, t) - 1)
                l = IIf(IsNumeric(l_), val(l_), 0)
                u = IIf(IsNumeric(u_), val(u_), 0)
                
                For j = l To u:   collect.Add j:   Next j
                GoTo nexti
            End If
            
            ' Unsupported string
            Call Err.Raise(9002, "DecodeIndexStringToArray", "argument string is unsupported.")
nexti:
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

Private Function RedimCollectionToArray2D(coll, lbound1 As Long, ubound1 As Long, lbound2 As Long, ubound2 As Long, Optional blnk As Variant = "", Optional isDirectionColumn As Boolean = False) As Variant()
    Dim rearr As Variant, c As Variant, i As Long, j As Long
    
    ReDim rearr(lbound1 To ubound1, lbound2 To ubound2)
    SetValueToAllElement rearr, blnk
    i = lbound1: j = lbound2
    For Each c In coll
        rearr(i, j) = c
        If isDirectionColumn = True Then: If j + 1 > ubound2 Then i = i + 1: j = lbound2 Else j = j + 1
        If isDirectionColumn = False Then: If i + 1 > ubound1 Then j = j + 1: i = lbound1 Else i = i + 1
        If i > ubound1 Or j > ubound2 Then GoTo outLoop
    Next
outLoop:
        
    RedimCollectionToArray2D = rearr
End Function

Private Function SetValueToAllElement(ByRef arr As Variant, val As Variant) As Variant()
    Dim i As Long, j As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        For j = LBound(arr, 2) To UBound(arr, 2)
            arr(i, j) = val
        Next j
    Next i
End Function

Private Function GetElementsCollection(arr, Optional isDirectionColumn As Boolean = True) As Collection
    Dim col As Collection: Set col = New Collection
    Dim ele, i As Long, j As Long
    i = LBound(arr, 1): j = LBound(arr, 2)
    For Each ele In arr
        col.Add arr(i, j)
        If isDirectionColumn = True Then: If j + 1 > UBound(arr, 2) Then i = i + 1: j = LBound(arr, 2) Else j = j + 1
        If isDirectionColumn = False Then: If i + 1 > UBound(arr, 1) Then j = j + 1: i = LBound(arr, 1) Else i = i + 1
    Next
    
    Set GetElementsCollection = col
End Function

Private Function AddElementsCollection(ByRef col As Collection, arr, Optional isDirectionColumn As Boolean = True)
    Dim ele, i As Long, j As Long
    i = LBound(arr, 1): j = LBound(arr, 2)
    For Each ele In arr
        col.Add arr(i, j)
        If isDirectionColumn = True Then: If j + 1 > UBound(arr, 2) Then i = i + 1: j = LBound(arr, 2) Else j = j + 1
        If isDirectionColumn = False Then: If i + 1 > UBound(arr, 1) Then j = j + 1: i = LBound(arr, 1) Else i = i + 1
    Next
End Function

Private Function StackArray(arrays As Collection, isDirectionColumn As Boolean) As Variant
    Dim paramarr, ele, array1, rearr, cnt As Long, dimension As Long
    Dim eles As Collection: Set eles = New Collection
    
    dimension = IIf(isDirectionColumn, 1, 2)
    For Each paramarr In arrays
        cnt = IIf(UBound(paramarr, dimension) - LBound(paramarr, dimension) + 1 > cnt, UBound(paramarr, dimension) - LBound(paramarr, dimension) + 1, cnt)
    Next
    
    For Each paramarr In arrays
        If isDirectionColumn Then
            paramarr = RedimPreserve2(paramarr, LBound(paramarr, 1), LBound(paramarr, 1) + cnt - 1, LBound(paramarr, 2), UBound(paramarr, 2), CVErr(xlErrNA))
        Else
            paramarr = RedimPreserve2(paramarr, LBound(paramarr, 1), UBound(paramarr, 1), LBound(paramarr, 2), LBound(paramarr, 2) + cnt - 1, CVErr(xlErrNA))
        End If
        AddElementsCollection eles, paramarr, Not (isDirectionColumn)
    Next
    
    array1 = arrays(1)
    If isDirectionColumn Then
        rearr = RedimCollectionToArray2D(eles, LBound(array1, 1), LBound(array1, 1) + cnt - 1, LBound(array1, 2), eles.Count / cnt, , Not (isDirectionColumn))
    Else
        rearr = RedimCollectionToArray2D(eles, LBound(array1, 1), eles.Count / cnt, LBound(array1, 2), LBound(array1, 2) + cnt - 1, , Not (isDirectionColumn))
    End If
    StackArray = rearr
End Function

Private Function GetFromCollection(coll As Collection)
    Dim c, re, cnt As Long
    cnt = 1
    ReDim re(1 To coll.Count)
    For Each c In coll
        re(cnt) = c
        cnt = cnt + 1
    Next
    GetFromCollection = re
End Function

Private Function GetSplitStringCollection(txt As String, delimiters, ignore_empty As Boolean, match_mode As Long, Optional instance_num As Long = 1000) As Collection
    Dim col As Collection:    Set col = New Collection
    Dim ind, extractText, removedText, i As Long
    i = 1
    Do While 1
        ind = GetLatestDelimiterIndex(txt, delimiters, extractText, removedText, match_mode, instance_num > 0)
        If ind = 0 Then GoTo outLoop
        If ignore_empty And extractText = "" Then Else col.Add extractText
        txt = removedText
        If Abs(instance_num) = i Then GoTo outLoop
        i = i + 1
    Loop
outLoop:
    If ignore_empty And txt = "" Then Else col.Add txt
    Set GetSplitStringCollection = col
End Function

Private Function GetLatestDelimiterIndex(text As String, delimiters, ByRef extractText, ByRef removedText, match_mode As Long, Optional searchFront As Boolean = True) As Long
    Dim delim, index As Long, ind As Long, f As Boolean, ff As Boolean, delimiter
    index = IIf(searchFront, 1000000, 0)
    f = False
    If match_mode = 1 Then text = LCase(text): delim = LCase(delim)
    For Each delim In delimiters
        ind = IIf(searchFront, InStr(text, delim), InStrRev(text, delim))
        f = ind > 0 Or f
        ff = IIf(searchFront, (ind > 0 And ind <= index), (ind > 0 And ind >= index))
        If ff Then
            index = ind
            delimiter = delim
        End If
    Next
    index = IIf(f, index, 0)
    If index = 0 Then Exit Function
    extractText = IIf(searchFront, Mid(text, 1, index - 1), Mid(text, index + Len(delimiter)))
    removedText = IIf(searchFront, Mid(text, index + Len(delimiter)), Mid(text, 1, index - 1))
    GetLatestDelimiterIndex = index
End Function

Private Function GetMaxCount(ByVal col As Collection) As Long
    Dim cmax As Long, co
    For Each co In col
        cmax = IIf(co.Count > cmax, co.Count, cmax)
    Next
    GetMaxCount = cmax
End Function

Private Function GetArraysParameters(ByVal arrays, ByRef lb1 As Long, ByRef ub1 As Long, ByRef cnt1 As Long, ByRef lb2 As Long, ByRef ub2 As Long, ByRef cnt2 As Long)
    Dim arr, rnk As Long

    lb1 = 100: lb2 = 100
    For Each arr In arrays
        rnk = Rank(arr)
        If rnk > 0 Then
            lb1 = IIf(lb1 > LBound(arr, 1), LBound(arr, 1), lb1)
            ub1 = IIf(ub1 < UBound(arr, 1), UBound(arr, 1), ub1)
            cnt1 = ub1 - lb1 + 1
        End If
        If rnk > 1 Then
            lb2 = IIf(lb2 > LBound(arr, 2), LBound(arr, 2), lb2)
            ub2 = IIf(ub2 < UBound(arr, 2), UBound(arr, 2), ub2)
            cnt2 = ub2 - lb2 + 1
        End If
    Next
End Function

Private Sub GetArrayParameters(ByVal arr, ByRef rnk As Long, ByRef lb1 As Long, ByRef ub1 As Long, ByRef cnt1 As Long, ByRef lb2 As Long, ByRef ub2 As Long, ByRef cnt2 As Long)
'    Dim i as long, j as long, rnk As Long, lb1 As Long, ub1 As Long, lb2 As Long, ub2 As Long, cnt1 As Long, cnt2 As Long
    rnk = Rank(arr)
    If rnk = 0 Then _
        Exit Sub
    If rnk > 0 Then _
        lb1 = LBound(arr, 1): ub1 = UBound(arr, 1): cnt1 = ub1 - lb1 + 1
    If rnk > 1 Then _
        lb2 = LBound(arr, 2): ub2 = UBound(arr, 2): cnt2 = ub2 - lb2 + 1
    If rnk > 2 Then _
        Err.Raise 9100, "ArrayExCore.GetArrayParameters", "array's dimention is over 2."
End Sub



