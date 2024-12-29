VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ------------------------------------------------------------------------------------------------------------
' Project                   : https://github.com/yyukki5/ArrayEx
' Version                   : 3.0.0
' Dependencies              : Lambda
' ------------------------------------------------------------------------------------------------------------

Private Const ERR_NUM_UNSUPPORT_RANK As Long = 9001
Private Const ERR_NUM_UNSUPPORT_ARGUMENT As Long = 9002
Private Const ERR_DESC_UNSUPPORT_RANK As String = "unsupported array rank"
Private Const ERR_DESC_UNSUPPORT_ARGUMENT As String = "unsupported argument"

Private value_ As Variant
Private rank_ As Long
Private lb1_ As Long
Private ub1_ As Long
Private lb2_ As Long
Private ub2_ As Long

Public Enum ArrayExOrientations
    ArrayExVertical = 1
    ArrayExHorizontal = 2
End Enum

Public Enum ArrayExRanks
    ArrayExRank0 = 0
    ArrayExRank1 = 1
    ArrayExRank2 = 2
End Enum

Public Function Initialize(ByVal val As Variant) As ArrayEx
    
    value_ = val
    rank_ = GetRank(value_)
    If rank_ > 0 Then lb1_ = LBound(value_, 1): ub1_ = UBound(value_, 1)
    If rank_ > 1 Then lb2_ = LBound(value_, 2): ub2_ = UBound(value_, 2)
    If rank_ > 2 Then Call Err.Raise(ERR_NUM_UNSUPPORT_RANK, "ArrayEx.Initialize()", ERR_DESC_UNSUPPORT_RANK)
    
    Set Initialize = Me
End Function

Public Function Create(val As Variant) As ArrayEx
Attribute Create.VB_Description = "Return new instance of ArrEx"
Attribute Create.VB_UserMemId = 0
    Set Create = New ArrayEx
    Call Create.Initialize(val)
End Function

Public Property Get Value() As Variant
    Value = value_
End Property
Public Property Get Rank() As Long
    Rank = rank_
End Property
Public Property Get UpperBound(Optional d As Long = 1) As Long
    UpperBound = IIf(d = 1, ub1_, ub2_)
End Property
Public Property Get LowerBound(Optional d As Long = 1) As Long
    LowerBound = IIf(d = 1, lb1_, lb2_)
End Property
Public Property Get Count(Optional d As Long = 1) As Long
    If IsNull(value_) Then Count = 0: Exit Property
    Count = IIf(d = 1, (ub1_ - lb1_ + 1), (ub2_ - lb2_ + 1))
End Property



Public Function Equals(compared_value As Variant) As Boolean
    Dim i As Long, j As Long, val As Variant
    
    If IsObject(compared_value) Then
        If TypeName(compared_value) = "ArrayEx" Then
            val = compared_value.Value
        Else
            Call Err.Raise(ERR_NUM_UNSUPPORT_ARGUMENT, "ArrayEx.Equals()", ERR_DESC_UNSUPPORT_ARGUMENT & " (" & TypeName(compared_value) & ")")
        End If
    Else
        val = compared_value
    End If
    
    If rank_ = 0 Then
        If value_ <> val Then Exit Function
    
    ElseIf rank_ = 1 Then
        If IsArray(val) = False Then Exit Function
        If (lb1_ = LBound(val, 1) And ub1_ = UBound(val, 1)) = False Then Exit Function
        For i = LBound(val, 1) To UBound(val, 1)
            If value_(i) <> val(i) Then Exit Function
        Next i
        
    ElseIf rank_ = 2 Then
        If IsArray(val) = False Then Exit Function
        If (lb1_ = LBound(val, 1) And ub1_ = UBound(val, 1) And lb2_ = LBound(val, 2) And ub2_ = UBound(val, 2)) = False Then Exit Function
        For i = LBound(val, 1) To UBound(val, 1)
            For j = LBound(val, 2) To UBound(val, 2)
                If value_(i, j) <> val(i, j) Then Exit Function
            Next j
        Next i
    
    End If

    Equals = True
End Function

Public Function RedimPreserve(array_rank As ArrayExRanks, _
    Optional lower_bound1 As Variant, _
    Optional upper_bound1 As Variant, _
    Optional lower_bound2 As Variant, _
    Optional upper_bound2 As Variant, _
    Optional blank_value As Variant) As ArrayEx
    Dim i As Long, j As Long, arr, cnt1 As Long, cnt2 As Long, l1 As Long, l2 As Long, u1 As Long, u2 As Long
    
    If array_rank = ArrayExRank0 Then
        If (IsMissing(lower_bound1) And IsMissing(upper_bound1) And IsMissing(lower_bound2) And IsMissing(upper_bound2)) = False Then Call Err.Raise(ERR_NUM_UNSUPPORT_ARGUMENT, "ArrayEx.RedimPreserve()", ERR_DESC_UNSUPPORT_ARGUMENT)
    End If
    If array_rank > ArrayExRank0 Then
        If IsMissing(lower_bound1) = False And IsMissing(upper_bound1) = True Then Call Err.Raise(ERR_NUM_UNSUPPORT_ARGUMENT, "ArrayEx.RedimPreserve()", ERR_DESC_UNSUPPORT_ARGUMENT & vbCrLf & "missing upper_bound1")
        If IsMissing(lower_bound1) = True And IsMissing(upper_bound1) = False Then Call Err.Raise(ERR_NUM_UNSUPPORT_ARGUMENT, "ArrayEx.RedimPreserve()", ERR_DESC_UNSUPPORT_ARGUMENT & vbCrLf & "missing lower_bound1")
    End If
    If array_rank > ArrayExRank1 Then
        If IsMissing(lower_bound2) = False And IsMissing(upper_bound2) = True Then Call Err.Raise(ERR_NUM_UNSUPPORT_ARGUMENT, "ArrayEx.RedimPreserve()", ERR_DESC_UNSUPPORT_ARGUMENT & vbCrLf & "missing upper_bound2")
        If IsMissing(lower_bound2) = True And IsMissing(upper_bound2) = False Then Call Err.Raise(ERR_NUM_UNSUPPORT_ARGUMENT, "ArrayEx.RedimPreserve()", ERR_DESC_UNSUPPORT_ARGUMENT & vbCrLf & "missing lower_bound2")
    End If
    
    If IsMissing(lower_bound1) Then l1 = lb1_ Else l1 = lower_bound1
    If IsMissing(upper_bound1) Then u1 = ub1_ Else u1 = upper_bound1
    If IsMissing(lower_bound2) Then l2 = lb2_ Else l2 = lower_bound2
    If IsMissing(upper_bound2) Then u2 = ub2_ Else u2 = upper_bound2
    
    cnt1 = u1 - l1
    cnt2 = u2 - l2
    
    Select Case (rank_ * 10 + array_rank)
        Case 0
            arr = value_
        Case 1
            ReDim arr(l1 To u1)
            If IsMissing(blank_value) = False Then arr = SetValueToAllElementByRank(arr, 1, blank_value)
            arr(l1) = value_
        
        Case 2
            ReDim arr(l1 To u1, l2 To u2)
            If IsMissing(blank_value) = False Then arr = SetValueToAllElementByRank(arr, 2, blank_value)
            arr(l1, l2) = value_
        
        Case 10
            arr = value_(lb1_)
        
        Case 11
            ReDim arr(l1 To u1)
            u1 = IIf((ub1_ - lb1_ > cnt1), cnt1, ub1_ - lb1_)
            If IsMissing(blank_value) = False Then arr = SetValueToAllElementByRank(arr, 1, blank_value)
            For i = 0 To u1: arr(l1 + i) = value_(lb1_ + i):  Next i
        
        Case 12
            ReDim arr(l1 To u1, l2 To u2)
            u2 = IIf((ub1_ - lb1_) > cnt2, cnt2, ub1_ - lb1_)
            If IsMissing(blank_value) = False Then arr = SetValueToAllElementByRank(arr, 2, blank_value)
            For i = 0 To u2:    arr(l1, l2 + i) = value_(lb1_ + i): Next i
        
        Case 20
            arr = value_(lb1_, lb2_)
        
        Case 21
            ReDim arr(l2 To u2)
            u2 = IIf(ub2_ - lb2_ > cnt2, cnt2, ub2_ - lb2_)
            If IsMissing(blank_value) = False Then arr = SetValueToAllElementByRank(arr, 1, blank_value)
            For i = 0 To u2: arr(l2 + i) = value_(lb1_, lb2_ + i):  Next i
        
        Case 22
            ReDim arr(l1 To u1, l2 To u2)
            u1 = IIf((ub1_ - lb1_) > cnt1, cnt1, ub1_ - lb1_)
            u2 = IIf((ub2_ - lb2_) > cnt2, cnt2, ub2_ - lb2_)
            
            If IsMissing(blank_value) = False Then arr = SetValueToAllElementByRank(arr, 2, blank_value)
            
            For i = 0 To u1
                For j = 0 To u2
                    arr(l1 + i, l2 + j) = value_(lb1_ + i, lb2_ + j)
                Next j
            Next i
    
    End Select

    Set RedimPreserve = Create(arr)
End Function

Public Function ShiftIndex(add_index1 As Long, Optional add_index2 As Long) As ArrayEx
    Dim rearr, i As Long, j As Long
    
    If rank_ = 0 Then Call Err.Raise(ERR_NUM_UNSUPPORT_RANK, "ArrayEx.ShiftIndex()", "Need dimension more 0")
    If rank_ > 0 Then If lb1_ + add_index1 < 0 Then Call Err.Raise(ERR_NUM_UNSUPPORT_ARGUMENT, "ArrayEx.ShiftIndex()", "Need add_index1 is 0 or more")
    If rank_ > 1 Then If lb2_ + add_index2 < 0 Then Call Err.Raise(ERR_NUM_UNSUPPORT_ARGUMENT, "ArrayEx.ShiftIndex()", "Need add_index2 is 0 or more")

    If rank_ = 1 Then
        ReDim rearr(lb1_ + add_index1 To ub1_ + add_index1)
        For i = lb1_ To ub1_
            rearr(i + add_index1) = value_(i)
        Next i
        
    ElseIf rank_ = 2 Then
        ReDim rearr(lb1_ + add_index1 To ub1_ + add_index1, lb2_ + add_index2 To ub2_ + add_index2)
        For i = lb1_ To ub1_
            For j = lb2_ To ub2_
                rearr(i + add_index1, j + add_index2) = value_(i, j)
            Next j
        Next i
    End If
    
    Set ShiftIndex = Create(rearr)
End Function

Public Function SelectColumns(ParamArray any_indexes()) As ArrayEx
    Dim rearr, i As Long, j As Long, col As New Collection, ind As Long, v As Variant, ub1 As Long, ub2 As Long
    
    If IsEmpty(any_indexes) Then Set SelectColumns = Create(Null): Exit Function
    
    If IsObject(any_indexes(0)) And TypeName(any_indexes(0)) = "Collection" Then
        Set col = any_indexes(0)
    
    ElseIf IsArray(any_indexes(0)) Then
        any_indexes = any_indexes(0)
        For Each v In any_indexes: Call col.Add(CLng(v)): Next
    
    Else
        For Each v In any_indexes: Call col.Add(CLng(v)): Next
        
    End If
    
    If col.Count = 0 Then Set SelectColumns = Create(Null): Exit Function
    
    
    If rank_ = 0 Then
        rearr = value_
    
    ElseIf rank_ = 1 Then
        ub1 = lb1_ + col.Count - 1
        ReDim rearr(lb1_ To ub1)
        For j = lb1_ To ub1
            ind = j - lb1_ + 1
            rearr(j) = value_(col(ind))
        Next j
    
    ElseIf rank_ = 2 Then
        ub2 = lb2_ + col.Count - 1
        ReDim rearr(lb1_ To ub1_, lb2_ To ub2)
        For i = lb1_ To ub1_
            For j = lb2_ To ub2
                ind = j - lb2_ + 1
                rearr(i, j) = value_(i, col(ind))
            Next j
        Next
    End If

    Set SelectColumns = Create(rearr)
End Function

Public Function SelectRows(ParamArray row_indexes()) As ArrayEx
    Dim rearr, i As Long, j As Long, col As New Collection, ind As Long, v As Variant, ub1 As Long
    
    If IsEmpty(row_indexes) Then Set SelectRows = Create(Null): Exit Function
    
    If IsObject(row_indexes(0)) And TypeName(row_indexes(0)) = "Collection" Then
        Set col = row_indexes(0)
    
    ElseIf IsArray(row_indexes(0)) Then
        row_indexes = row_indexes(0)
        For Each v In row_indexes: Call col.Add(CLng(v)): Next
    
    Else
        For Each v In row_indexes: Call col.Add(CLng(v)): Next
        
    End If
    
    If col.Count = 0 Then Set SelectRows = Create(Null): Exit Function
        
    
    If (rank_ = 2) = False Then
        rearr = value_
      
    Else
        ub1 = lb1_ + col.Count - 1
        ReDim rearr(lb1_ To ub1, lb2_ To ub2_)
        For i = lb1_ To ub1
            ind = i - lb1_ + 1
            For j = lb2_ To ub2_
                rearr(i, j) = value_(col(ind), j)
            Next j
        Next
    End If
        
    Set SelectRows = Create(rearr)
End Function


Public Function WhereBy(lambda_string As String) As ArrayEx
    Dim col As New Collection, i As Long, j As Long
    
    If rank_ = 0 Then
        Set WhereBy = Create(value_)
        
    ElseIf rank_ = 1 Then
        For i = lb1_ To ub1_
            If Lambda(value_(i), lambda_string).Evaluate Then Call col.Add(i)
        Next i
        Set WhereBy = SelectColumns(col)
            
    ElseIf rank_ = 2 Then
        For i = lb1_ To ub1_
            If Lambda(GetRow(value_, i), lambda_string).Evaluate Then Call col.Add(i)
        Next i
        Set WhereBy = SelectRows(col)

    End If
    
End Function

Public Function SelectBy(lambda_string As String) As ArrayEx
    Dim col As New Collection, i As Long, j As Long, re As Variant
    
    If rank_ = 0 Then
        re = Lambda(value_, lambda_string).Evaluate
        
    ElseIf rank_ = 1 Then
        ReDim re(lb1_ To ub1_)
        For i = lb1_ To ub1_
            re(i) = Lambda(value_(i), lambda_string).Evaluate
        Next i
            
    ElseIf rank_ = 2 Then
        ReDim re(lb1_ To ub1_)
        For i = lb1_ To ub1_
            re(i) = Lambda(GetRow(value_, i), lambda_string).Evaluate
        Next i

    End If
    
    Set SelectBy = Create(re)
End Function

Public Function Skip(skip_index As Long, Optional column_skip_index As Long = -1) As ArrayEx
    Dim col As New Collection, col1 As New Collection, i As Long, j As Long, ub1 As Long, ub2 As Long
    
    If rank_ = 0 Then
        Set Skip = Create(value_)
        
    ElseIf rank_ = 1 Then
        For i = lb1_ + skip_index To ub1_:  Call col.Add(i): Next
        Set Skip = SelectColumns(col)

    ElseIf rank_ = 2 Then
        If column_skip_index = -1 Then
            For i = lb1_ + skip_index To ub1_:  Call col.Add(i): Next
            Set Skip = SelectRows(col)
        Else
            For i = lb1_ + skip_index To ub1_:  Call col.Add(i): Next
            For j = lb2_ + column_skip_index To ub2_:  Call col1.Add(j): Next
            Set Skip = SelectRows(col).SelectColumns(col1)
        End If
    End If
    
End Function

Public Function Take(take_index As Long, Optional column_take_index As Long = -1) As ArrayEx
    Dim col As New Collection, col1 As New Collection, i As Long, j As Long, ub1 As Long, ub2 As Long
    
    If rank_ = 0 Then
        Set Take = Create(value_)
        
    ElseIf rank_ = 1 Then
        ub1 = IIf(lb1_ + take_index - 1 < ub1_, lb1_ + take_index - 1, ub1_)
    
        For i = lb1_ To ub1:  Call col.Add(i): Next
        Set Take = SelectColumns(col)

    ElseIf rank_ = 2 Then
        If column_take_index = -1 Then
            ub1 = IIf(lb1_ + take_index - 1 < ub1_, lb1_ + take_index - 1, ub1_)
            For i = lb1_ To ub1:  Call col.Add(i): Next
            Set Take = SelectRows(col)
            
        Else
            ub1 = IIf(lb1_ + take_index - 1 < ub1_, lb1_ + take_index - 1, ub1_)
            ub2 = IIf(lb2_ + column_take_index - 1 < ub2_, lb2_ + column_take_index - 1, ub2_)
            For i = lb1_ To ub1:  Call col.Add(i): Next
            For j = lb2_ To ub2:  Call col1.Add(j): Next
            Set Take = SelectRows(col).SelectColumns(col1)
        
        End If
    End If

End Function

Public Function OrderBy(Optional column_index As Variant) As ArrayEx
    If rank_ = 0 Then
        Set OrderBy = Create(value_)
    ElseIf rank_ = 1 Then
        Set OrderBy = Create(Sort1Dim(value_, True))
    ElseIf rank_ = 2 Then
        If IsMissing(column_index) Then Call Err.Raise(ERR_NUM_UNSUPPORT_ARGUMENT, "ArrayEx.OrderBy()", "missing column_index")
        Set OrderBy = Create(Sort2Dim(value_, CLng(column_index), True))
    End If
End Function

Public Function OrderByDescending(Optional column_index As Long) As ArrayEx
    If rank_ = 0 Then
        Set OrderByDescending = Create(value_)
    ElseIf rank_ = 1 Then
        Set OrderByDescending = Create(Sort1Dim(value_, False))
    ElseIf rank_ = 2 Then
        If IsMissing(column_index) Then Call Err.Raise(ERR_NUM_UNSUPPORT_ARGUMENT, "ArrayEx.OrderByDescending()", "missing column_index")
        Set OrderByDescending = Create(Sort2Dim(value_, CLng(column_index), False))
    End If
End Function

Public Function Distinct() As ArrayEx
    Dim i As Long, any_indexes As New Collection, vals As New Collection
    
    If rank_ = 0 Then
        Set Distinct = Create(value_)
    
    ElseIf rank_ = 1 Then
        For i = lb1_ To ub1_
            If CheckExist(vals, value_(i)) = False Then
                Call vals.Add(value_(i))
                Call any_indexes.Add(i)
            End If
        Next i
        Set Distinct = SelectColumns(any_indexes)
            
    ElseIf rank_ = 2 Then
        Dim aex As New ArrayEx
        For i = lb1_ To ub1_
            Set aex = SelectRows(i)
            If CheckExist(vals, aex) = False Then
                Call vals.Add(aex)
                Call any_indexes.Add(i)
            End If
        Next i
        Set Distinct = SelectRows(any_indexes)
            
    End If
    
End Function

Public Function DistinctBy(column_index As Long) As ArrayEx
    Dim i As Long, row_indexes As New Collection, vals As New Collection
    
    If rank_ <> 2 Then Call Err.Raise(ERR_NUM_UNSUPPORT_RANK, "ArrayEx.DistinctBy()", ERR_DESC_UNSUPPORT_RANK)
    
    For i = lb1_ To ub1_
        If CheckExist(vals, value_(i, column_index)) = False Then
            Call vals.Add(value_(i, column_index))
            Call row_indexes.Add(i)
        End If
    Next i
        
    Set DistinctBy = SelectRows(row_indexes)
End Function

Public Function Transpose() As ArrayEx
    Dim i As Long, j As Long, re As Variant
    
    If rank_ = 0 Then
        re = value_
    
    ElseIf rank_ = 1 Then
        ReDim re(lb1_ To ub1_, 0 To 0)
        For i = lb1_ To ub1_
            re(i, 0) = value_(i)
        Next i

    ElseIf rank_ = 2 Then
        ReDim re(lb2_ To ub2_, lb1_ To ub1_)
        For i = lb1_ To ub1_
            For j = lb2_ To ub2_
                re(j, i) = value_(i, j)
            Next j
        Next i
    End If

    Set Transpose = Create(re)
End Function

Public Function VerticalStack(arr As Variant, Optional blank_value As Variant = Null) As ArrayEx
    Set VerticalStack = Create(StackArray(ArrayExVertical, arr, blank_value))
End Function

Public Function HorizontalStack(arr As Variant, Optional blank_value As Variant = Null) As ArrayEx
    Set HorizontalStack = Create(StackArray(ArrayExHorizontal, arr, blank_value))
End Function

' To
Public Function ConvertElementsTo(var_type As VbVarType) As ArrayEx
    Dim a As Variant, value1 As Variant, col As New Collection
    
    If rank_ = 0 Then value1 = Array(value_) Else value1 = value_
    
    For Each a In value1
        Select Case var_type
            Case vbBoolean: Call col.Add(CBool(a))
            Case vbByte: Call col.Add(CByte(a))
            Case vbCurrency: Call col.Add(CCur(a))
            Case vbDate: Call col.Add(CDate(a))
            Case vbDecimal: Call col.Add(CDec(a))
            Case vbDouble: Call col.Add(CDbl(a))
            Case vbInteger: Call col.Add(CInt(a))
            Case vbLong: Call col.Add(CLng(a))
            Case vbLongLong: Call col.Add(CLngLng(a))
            Case vbSingle: Call col.Add(CSng(a))
            Case vbString:  Call col.Add(CStr(a))
            Case Else: Call Err.Raise(ERR_NUM_UNSUPPORT_ARGUMENT, "ArrayEx.ConvertElementsTo()", ERR_DESC_UNSUPPORT_ARGUMENT)
        End Select
    Next
    
    Set ConvertElementsTo = Create(CollectionToArray(col, rank_))
End Function

Public Function ToCollection(Optional split_orientation As ArrayExOrientations) As Collection
    Dim col As New Collection, a As Variant, value1 As Variant, i As Long
    
    If rank_ = 0 Then value1 = Array(value_) Else value1 = value_
    
    If rank_ = 0 Then
        Call col.Add(value1)
    ElseIf rank_ = 1 Then
        For Each a In value1:  Call col.Add(a): Next
    
    ElseIf rank_ = 2 Then
        If split_orientation = 0 Then
            For Each a In value1:  Call col.Add(a): Next
        
        ElseIf split_orientation = ArrayExHorizontal Then
            For i = lb1_ To ub1_
                Call col.Add(GetRow(value1, i), CStr(i))
            Next i
    
        ElseIf split_orientation = ArrayExVertical Then
            For i = lb2_ To ub2_
                Call col.Add(GetColumn(value1, i), CStr(i))
            Next i
        Else
        End If
    
    End If
    
    Set ToCollection = col
End Function

Public Function ToString(Optional column_delimiter As String = ",", Optional row_delimiter As String = ";") As String
    Dim s As String, s1 As String, i As Long, j As Long
    
    If rank_ = 0 Then
        If IsNull(value_) = False Then s = CStr(value_)
    
    ElseIf rank_ = 1 Then
        For i = lb1_ To ub1_
            If IsNull(value_(i)) Then s = s & column_delimiter Else s = s & column_delimiter & CStr(value_(i))
        Next i
        s = Strings.Mid(s, Len(column_delimiter) + 1)
    
    ElseIf rank_ = 2 Then
        For i = lb1_ To ub1_
            For j = lb2_ To ub2_
                If IsNull(value_(i, j)) Then s = s & column_delimiter Else s = s & CStr(value_(i, j)) & column_delimiter
            Next j
            s = Strings.Left(s, Len(s) - Len(column_delimiter)) & row_delimiter
        Next i
        s = Strings.Left(s, Len(s) - Len(row_delimiter))
    End If
    
    ToString = "" & s & ""
End Function


Public Function XLookUp(array_2d As Variant, result_col As Long, find_col As Long, Optional not_found As Variant = Null) As ArrayEx
    Dim v As Variant, col As New Collection, i As Long, re As Variant
    
    For Each v In value_
        re = not_found
        For i = LBound(array_2d, 1) To UBound(array_2d, 1) Step 1
            If array_2d(i, find_col) = v Then re = array_2d(i, result_col): Exit For
        Next i
        Call col.Add(re)
    Next
    
    Set XLookUp = Create(CollectionToArray(col, rank_))
End Function

' Join array
Public Function InnerJoin(arr As Variant, find_index As Long, joined_index As Long) As ArrayEx
    Dim recol As New Collection, ky As Variant, i As Long, j As Long, dict As Object
    
    
    Set dict = CreateObject("Scripting.Dictionary")
    For i = LBound(arr, 1) To UBound(arr, 1)
        ky = arr(i, joined_index)
        If dict.Exists(ky) = False Then Call dict.Add(ky, GetRow(arr, i))
    Next i
        
    For i = lb1_ To ub1_
        ky = value_(i, find_index)
        If dict.Exists(ky) Then
            Call recol.Add(JoinArrays(GetRow(value_, i), dict(ky)))
        End If
    Next i

    Set InnerJoin = Create(CollectedRowArrayToArray(recol))
End Function

Public Function LeftJoin(arr As Variant, find_index As Long, joined_index As Long) As ArrayEx
    Dim recol As New Collection, ky As Variant, i As Long, j As Long, dict As Object, tmp, ub As Long
    
    Set dict = CreateObject("Scripting.Dictionary")
    For i = LBound(arr, 1) To UBound(arr, 1)
        ky = arr(i, joined_index)
        If dict.Exists(ky) = False Then Call dict.Add(ky, GetRow(arr, i))
    Next i

    For i = lb1_ To ub1_
        ky = value_(i, find_index)
        If dict.Exists(ky) Then
            Call recol.Add(JoinArrays(GetRow(value_, i), dict(ky)))
        Else
            Call recol.Add(JoinArrays(GetRow(value_, i), CreateNullArray(LBound(arr, 2), UBound(arr, 2))))
        End If
    Next i

    Set LeftJoin = Create(CollectedRowArrayToArray(recol))
End Function

Public Function FullOuterJoin(arr As Variant, find_index As Long, joined_index As Long) As ArrayEx
    Dim recol As New Collection, ky As Variant, i As Long, j As Long, dict_l As Object, dict_r As Object, tmp, ub As Long
    
    Set dict_l = CreateObject("Scripting.Dictionary")
    For i = LBound(value_, 1) To UBound(value_, 1)
        ky = value_(i, find_index)
        If dict_l.Exists(ky) = False Then Call dict_l.Add(ky, GetRow(value_, i))
    Next i
        
    Set dict_r = CreateObject("Scripting.Dictionary")
    For i = LBound(arr, 1) To UBound(arr, 1)
        ky = arr(i, joined_index)
        If dict_r.Exists(ky) = False Then Call dict_r.Add(ky, GetRow(arr, i))
    Next i

    For Each ky In dict_l.Keys
        If dict_r.Exists(ky) Then
            Call recol.Add(JoinArrays(dict_l(ky), dict_r(ky)))
            Call dict_r.Remove(ky)
        Else
            Call recol.Add(JoinArrays(dict_l(ky), CreateNullArray(LBound(arr, 2), UBound(arr, 2))))
        End If
    Next
    
    For Each ky In dict_r.Keys
        Call recol.Add(JoinArrays(CreateNullArray(LBound(arr, 2), UBound(arr, 2)), dict_r(ky)))
    Next

    Set FullOuterJoin = Create(CollectedRowArrayToArray(recol))
End Function

Public Function CrossJoin(arr As Variant) As ArrayEx
    Dim recol As New Collection, i As Long, j As Long
    
    For i = lb1_ To ub1_
        For j = LBound(arr) To UBound(arr)
            Call recol.Add(JoinArrays(GetRow(value_, i), GetRow(arr, j)))
        Next j
    Next i
    
    Set CrossJoin = Create(CollectedRowArrayToArray(recol))
End Function


' Strings
' Do for each element in value_
Public Function Format(Optional format_text, Optional first_day_of_week As VbDayOfWeek = vbSunday, Optional first_week_of_year As VbFirstWeekOfYear = vbFirstJan1) As ArrayEx
    Dim a As Variant, col As New Collection
    
    For Each a In value_
        Call col.Add(Strings.Format(a, format_text, first_day_of_week, first_week_of_year))
    Next

    Set Format = Create(CollectionToArray(col, rank_))
End Function
Public Function Replace(find_string As String, replace_string As String) As ArrayEx
    Dim a As Variant, col As New Collection
    
    For Each a In value_
        Call col.Add(Strings.Replace(a, find_string, replace_string))
    Next

    Set Replace = Create(CollectionToArray(col, rank_))
End Function
Public Function Left(string_length As Long) As ArrayEx
    Dim a As Variant, col As New Collection
    
    For Each a In value_
        Call col.Add(Strings.Left(a, string_length))
    Next

    Set Left = Create(CollectionToArray(col, rank_))
End Function
Public Function Right(string_length As Long) As ArrayEx
    Dim a As Variant, col As New Collection
    
    For Each a In value_
        Call col.Add(Strings.Right(a, string_length))
    Next

    Set Right = Create(CollectionToArray(col, rank_))
End Function
Public Function Mid(start_index As Long, Optional length_index) As ArrayEx
    Dim a As Variant, col As New Collection
    
    For Each a In value_
        If IsMissing(length_index) Then
            Call col.Add(Strings.Mid(a, start_index))
        Else
            Call col.Add(Strings.Mid(a, start_index, length_index))
        End If
    Next

    Set Mid = Create(CollectionToArray(col, rank_))
End Function


' Processing the + operator
Public Function Plus(val As Variant) As ArrayEx
    Dim a As Variant, col As New Collection
    
    On Error GoTo errl
    For Each a In value_
        Call col.Add(a + val)
    Next

    Set Plus = Create(CollectionToArray(col, rank_))
    Exit Function
errl:
    Call Err.Raise(Err.Number, Err.Source & vbCrLf & "at ArrayEx.Plus()", Err.Description)
End Function


' Developing tools
Public Function DebugPrint(Optional column_delimiter As String = ",") As ArrayEx
    Dim i As Long, j As Long, txt As String
    If rank_ = 0 Then
        Debug.Print value_
        
    ElseIf rank_ = 1 Then
        For j = lb1_ To ub1_
            txt = txt & column_delimiter & value_(j)
        Next j
        Debug.Print Strings.Mid(txt, Len(column_delimiter) + 1)
        
    ElseIf rank_ = 2 Then
        For i = lb1_ To ub1_
            For j = lb2_ To ub2_
                txt = txt & column_delimiter & value_(i, j)
            Next j
            Debug.Print Strings.Mid(txt, Len(column_delimiter) + 1)
            txt = ""
        Next i
    
    Else
    End If
    Debug.Print ""
    Set DebugPrint = Me
End Function


' private functions -------------------------------------------------------------------

Private Function Compare(compared_value As Variant, compare_operator As aexOperators, Optional compare_value As Variant = Null)
    If IsNull(compare_value) Then
    
    Else
        Select Case compare_operator
            Case aexEqual
                Compare = (compare_value = compared_value)
            Case aexGreaterThan
                Compare = (compared_value > compare_value)
            Case aexGreaterThanOrEqualTo
                Compare = (compared_value >= compare_value)
            Case aexLessThan
                Compare = (compared_value < compare_value)
            Case aexLessThanOrEqualTo
                Compare = (compared_value <= compare_value)
            
            Case aexIs
                Compare = (compared_value Is compare_value)
            Case aexIsNot
                Compare = (compared_value Is Not compare_value)
        End Select
    End If
End Function

Private Function CollectionToArray(col As Collection, rnk As ArrayExRanks, Optional l1 As Long, Optional u1 As Long, Optional l2 As Long, Optional u2 As Long)
    Dim i As Long, j As Long, arr, cnt
    
    If col.Count = 0 Then CollectionToArray = Array(): Exit Function
    
    Select Case rnk
        Case 0
            arr = col(1)
        Case 1
            If l1 = 0 Then l1 = lb1_
            ReDim arr(l1 To l1 + col.Count - 1)
            For i = l1 To l1 + col.Count - 1
                cnt = cnt + 1
                If IsObject(col(cnt)) Then Set arr(i) = col(cnt) Else arr(i) = col(cnt)
            Next i
        Case 2
            If l1 = 0 Then l1 = LBound(value_, 1)
            If u1 = 0 Then u1 = UBound(value_, 1)
            If l2 = 0 Then l2 = LBound(value_, 2)
            If u2 = 0 Then u2 = UBound(value_, 2)
            
            ReDim arr(l1 To u1, l2 To u2)
            For j = l2 To u2
                For i = l1 To u1
                    cnt = cnt + 1
                    If IsObject(cnt) Then Set arr(i, j) = col(cnt) Else arr(i, j) = col(cnt)
                Next i
            Next j
    End Select
        
    CollectionToArray = arr
End Function

Private Function CollectedRowArrayToArray(col As Collection)
    Dim i As Long, j As Long, a As Variant, cnt As Long, re As Variant, ub1 As Long, ub2 As Long
    
    ub1 = lb1_ + col.Count - 1
    ub2 = lb2_ + (UBound(col(1)) - LBound(col(1)))
    ReDim re(lb1_ To ub1, lb2_ To ub2)
    
    i = lb1_ - 1
    For Each a In col
        i = i + 1
        For j = lb2_ To ub2
            re(i, j) = a((j - lb2_) + LBound(a))
        Next j
    Next
        
    CollectedRowArrayToArray = re
End Function

Private Function GetRank(arr As Variant)
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
    GetRank = (i - 1)
End Function

Private Function SetValueToAllElementByRank(arr As Variant, rnk As ArrayExRanks, val As Variant) As Variant()
    Dim i As Long, j As Long, rearr As Variant
        
    If rnk = ArrayExRank0 Then
        arr = val
    
    ElseIf rnk = ArrayExRank1 Then
        ReDim rearr(LBound(arr) To UBound(arr))
        For i = LBound(arr) To UBound(arr)
            rearr(i) = val
        Next i
    
    ElseIf rnk = ArrayExRank2 Then
        ReDim rearr(LBound(arr, 1) To UBound(arr, 1), LBound(arr, 2) To UBound(arr, 2))
        For i = LBound(arr, 1) To UBound(arr, 1)
            For j = LBound(arr, 2) To UBound(arr, 2)
                rearr(i, j) = val
            Next j
        Next i
    End If

    SetValueToAllElementByRank = rearr
End Function


Private Function SetValueToAllElement(arr As Variant, val As Variant) As Variant()
    Dim i As Long, j As Long, rearr
    
    ReDim rearr(LBound(arr, 1) To UBound(arr, 1), LBound(arr, 2) To UBound(arr, 2))
    
    For i = LBound(arr, 1) To UBound(arr, 1)
        For j = LBound(arr, 2) To UBound(arr, 2)
            rearr(i, j) = val
        Next j
    Next i

    SetValueToAllElement = rearr
End Function

Private Function StackArray(d As ArrayExOrientations, arr As Variant, blank_value As Variant) As Variant
    Dim rearr As Variant, col As New Collection, aex As ArrayEx
    Dim lb1 As Long, ub1 As Long, lb2 As Long, ub2 As Long, cnt As Long, i As Long, j As Long
    
    lb1 = lb1_
    lb2 = lb2_
    Call col.Add(Me)
    Call col.Add(Create(arr))
        
    If d = ArrayExVertical Then
        
        ' get max/min index
        lb1 = lb1_
        For Each aex In col
            Select Case aex.Rank
                Case 0
                    cnt = cnt + 1
                    lb2 = IIf(lb2 > 0, 0, lb2)
                    ub2 = IIf(ub2 < 0, 0, ub2)
                Case 1
                    cnt = cnt + 1
                    lb2 = IIf(lb2 > aex.LowerBound(1), aex.LowerBound(1), lb2)
                    ub2 = IIf(ub2 < aex.UpperBound(1), aex.UpperBound(1), ub2)
                Case 2
                    cnt = cnt + aex.Count(1)
                    lb2 = IIf(lb2 > aex.LowerBound(2), aex.LowerBound(2), lb2)
                    ub2 = IIf(ub2 < aex.UpperBound(2), aex.UpperBound(2), ub2)
                Case Else
            End Select
        Next
    
        ReDim rearr(lb1 To (lb1 + cnt - 1), lb2 To ub2)
        rearr = SetValueToAllElement(rearr, blank_value)
        
        ' substitude
        cnt = lb1
        For Each aex In col
            If aex.Rank = 0 Then
                rearr(cnt, lb2) = aex.Value
                cnt = cnt + 1
            ElseIf aex.Rank = 1 Then
                For j = aex.LowerBound(1) To aex.UpperBound(1)
                    rearr(cnt, j) = aex.Value(j)
                Next j
                cnt = cnt + 1
            ElseIf aex.Rank = 2 Then
                For i = aex.LowerBound(1) To aex.UpperBound(1)
                    For j = aex.LowerBound(2) To aex.UpperBound(2)
                        rearr(cnt, j) = aex.Value(i, j)
                    Next j
                    cnt = cnt + 1
                Next i
            Else
            End If
        Next
    
    ElseIf d = ArrayExHorizontal Then
        
        ' get max/min index
        lb2 = lb2_
        For Each aex In col
            Select Case aex.Rank
                Case 0
                    cnt = cnt + 1
                    lb1 = IIf(lb1 > 0, 0, lb1)
                    ub1 = IIf(ub1 < 0, 0, ub1)
                Case 1
                    cnt = cnt + aex.Count(1)
                    lb1 = IIf(lb1 > 0, 0, lb1)
                    ub1 = IIf(ub1 < 0, 0, ub1)
                Case 2
                    cnt = cnt + aex.Count(2)
                    lb1 = IIf(lb1 > aex.LowerBound, aex.LowerBound, lb1)
                    ub1 = IIf(ub1 < aex.UpperBound, aex.UpperBound, ub1)
                Case Else
            End Select
        Next
        
        ReDim rearr(lb1 To ub1, lb2 To (lb2 + cnt - 1))
        rearr = SetValueToAllElement(rearr, blank_value)
        
        ' substitude
        cnt = lb2
        For Each aex In col
            If aex.Rank = 0 Then
                rearr(lb1, cnt) = aex.Value
                cnt = cnt + 1
            ElseIf aex.Rank = 1 Then
                For j = aex.LowerBound(1) To aex.UpperBound(1)
                    rearr(lb1, cnt) = aex.Value(j)
                    cnt = cnt + 1
                Next j
            ElseIf aex.Rank = 2 Then
                For j = aex.LowerBound(2) To aex.UpperBound(2)
                    For i = aex.LowerBound(1) To aex.UpperBound(1)
                        rearr(i, cnt) = aex.Value(i, j)
                    Next i
                    cnt = cnt + 1
                Next j
            Else
            End If
        Next
    
    End If
            
    StackArray = rearr
End Function


Private Function Sort1Dim(ByVal arr, is_asc As Boolean) As Variant
    Call Sort1DimSub(arr, LBound(arr), UBound(arr), is_asc)
    Sort1Dim = arr
End Function
Private Function Sort2Dim(ByVal arr, col As Long, is_asc As Boolean) As Variant
    Call Sort2dimSub(arr, LBound(arr, 1), UBound(arr, 1), is_asc, col)
    Sort2Dim = arr
End Function

Private Function Sort1DimSub(ByRef arr As Variant, ByVal l As Long, ByVal u As Long, is_asc As Boolean)
    Dim i As Long, j As Long, p As Variant, tmp As Variant
    i = l
    j = u
    p = arr((l + u) \ 2)
    
    Do While i <= j
        If is_asc Then
            Do While arr(i) < p: i = i + 1: Loop
            Do While arr(j) > p: j = j - 1: Loop
        Else
            Do While arr(i) > p: i = i + 1: Loop
            Do While arr(j) < p: j = j - 1: Loop
        End If
        If i <= j Then
            tmp = arr(i)
            arr(i) = arr(j)
            arr(j) = tmp
            i = i + 1
            j = j - 1
        End If
    Loop
    If l < j Then Call Sort1DimSub(arr, l, j, is_asc)
    If i < u Then Call Sort1DimSub(arr, i, u, is_asc)
    Sort1DimSub = arr
End Function

Private Function Sort2dimSub(ByRef arr As Variant, ByVal l As Long, ByVal u As Long, ByVal is_asc As Boolean, ByVal col As Long)
    Dim i As Long, j As Long, k As Long, p As Variant, tmp As Variant
    i = l
    j = u
    p = arr((l + u) \ 2, col)
    
    Do While i <= j
        If is_asc Then
            Do While arr(i, col) < p: i = i + 1: Loop
            Do While arr(j, col) > p: j = j - 1: Loop
        Else
            Do While arr(i, col) > p: i = i + 1: Loop
            Do While arr(j, col) < p: j = j - 1: Loop
        End If
        If i <= j Then
            For k = LBound(arr, 2) To UBound(arr, 2)
                tmp = arr(i, k)
                arr(i, k) = arr(j, k)
                arr(j, k) = tmp
            Next k
            i = i + 1
            j = j - 1
        End If
    Loop
    If l < j Then Call Sort2dimSub(arr, l, j, is_asc, col)
    If i < u Then Call Sort2dimSub(arr, i, u, is_asc, col)
    Sort2dimSub = arr
End Function

Private Function CheckExist(col As Collection, val) As Boolean
    Dim v As Variant
    
    CheckExist = True
    For Each v In col
        If IsObject(v) Then
            If v.Equals(val) Then Exit Function
        Else
            If v = val Then Exit Function
        End If
    Next
    CheckExist = False

End Function


' Join Array
Private Function JoinArrays(arr1 As Variant, arr2 As Variant) As Variant
    Dim res() As Variant
    Dim i As Long, j As Long
    
    ReDim res(LBound(arr1) To UBound(arr1) + UBound(arr2) - LBound(arr2) + 1)
    For i = LBound(arr1) To UBound(arr1)
        res(i) = arr1(i)
    Next i
    For j = LBound(arr2) To UBound(arr2)
        res(UBound(arr1) + j - LBound(arr2) + 1) = arr2(j)
    Next j
    
    JoinArrays = res
End Function

Private Function CreateNullArray(lb As Long, ub As Long) As Variant
    Dim arr() As Variant, i As Long
    
    ReDim arr(lb To ub)
    For i = lb To ub
        arr(i) = Null
    Next i
    
    CreateNullArray = arr
End Function

Private Function GetRow(arr, ind)
    Dim re, i As Long
    ReDim re(LBound(arr, 2) To UBound(arr, 2))
    For i = LBound(arr, 2) To UBound(arr, 2): re(i) = arr(ind, i): Next i
    GetRow = re
End Function

Private Function GetColumn(arr, ind)
    Dim re, i As Long
    ReDim re(LBound(arr, 1) To UBound(arr, 1))
    For i = LBound(arr, 1) To UBound(arr, 1): re(i) = arr(i, ind): Next i
    GetColumn = re
End Function
