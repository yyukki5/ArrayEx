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

Private array_ As Variant

Public Function Init(array1d As Variant) As ArrayEx1
    If ArrayRank(array1d) <> 1 Then Exit Function
    array_ = array1d
    Set Init = Me
End Function

Public Function Create(Optional arr As Variant) As ArrayEx1
    If ArrayRank(arr) <> 1 Then Exit Function
    Dim rearr As New ArrayEx1
    Set Create = IIf(IsMissing(arr), rearr.Init(array_), rearr.Init(arr))
End Function

Public Property Get Value()
Attribute Value.VB_UserMemId = 0
    Value = array_
End Property

Public Property Let Value(v)
    array_ = v
End Property

Public Function ToCollection() As Collection
    Dim coll As New Collection
    Dim arr As New ArrayEx0
    
    Dim i As Long
    For i = Lb To Ub
        coll.Add arr.Create(array_(i))
    Next i

    Set ToCollection = coll
End Function

Public Function GetItem(index As Long) As ArrayEx0
    Dim rearr As New ArrayEx0
    Set GetItem = rearr.Init(array_(index))
End Function

Public Function Match(arg, Optional matchType As Long = 0) As Long
   Match = WorksheetFunction.Match(arg, array_, matchType)
End Function

Public Function DebugPrintAllItems() As ArrayEx1
    Debug.Print Join(array_, ",")
    Set DebugPrintAllItems = Me
End Function

Public Function DebugPrintEvaluated(formulaWithArgx As String) As ArrayEx1
    Dim i&, s$
    For i = Lb To Ub
        s = Replace(formulaWithArgx, "{x}", array_(i))
       Debug.Print Evaluate(s)
    Next i
    Set DebugPrintEvaluated = Me
End Function

Public Function DebugPrint(Optional pre$, Optional formatString$, Optional suf$) As ArrayEx1
    Dim i&
    For i = Lb To Ub
       Debug.Print pre & Format(array_(i), formatString) & suf
    Next i
    Set DebugPrint = Me
End Function

Public Function SelectEvaluated(formulaWithArgx As String) As ArrayEx1
Dim i As Long, rearr(), s As String
ReDim rearr(Lb To Ub)
For i = Lb To Ub
    s = Replace(formulaWithArgx, "{x}", array_(i))
   rearr(i) = Evaluate(s)
Next i
    
    Dim rearr0 As New ArrayEx1
    Set SelectEvaluated = rearr0.Init(rearr)
End Function

Public Function WhereEvaluated(formulaWithArgx As String) As ArrayEx1
Dim i As Long, rearr(), s As String
ReDim rearr(Lb To Ub)
Dim f As Boolean
Dim collect As New Collection
For i = Lb To Ub
    s = Replace(formulaWithArgx, "{x}", array_(i))
   f = Evaluate(s)
   If f Then collect.Add array_(i)
Next i
    
    Dim rearr1 As New ArrayEx1
    Set WhereEvaluated = rearr1.CreateFromCollection(collect)
End Function

Public Function CreateFromCollection(collect As Collection) As ArrayEx1
If ArrayRank(collect(1)) <> 0 Then Exit Function
    Dim rearr As New ArrayEx1
    Dim arr As Variant
    ReDim arr(1 To collect.Count)
    Dim c As Variant, i As Long:      i = 1
    For Each c In collect
            arr(i) = c
            i = i + 1
    Next c
 Set CreateFromCollection = rearr.Init(arr)
End Function



Public Property Get Ub()
    Ub = UBound(array_, 1)
End Property

Public Property Get Lb()
    Lb = LBound(array_, 1)
End Property

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
