VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayEx0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private array_ As Variant

Public Function Init(array0d) As ArrayEx0
    array_ = array0d
    Set Init = Me
End Function

Public Function Create(Optional arr As Variant) As ArrayEx0
    Dim rearr As New ArrayEx0
    Set Create = IIf(IsMissing(arr), rearr.Init(array_), rearr.Init(arr))
End Function

Public Property Get Value()
Attribute Value.VB_UserMemId = 0
    Value = array_
End Property

Public Function Replaced(before, after) As ArrayEx0
    ' array_ type : string only
    
    Dim rearr As New ArrayEx0
    Dim re
    re = Replace(array_, before, after)
    Set Replaced = rearr.Init(re)
End Function

Public Function ToString(Optional stringFormat As String) As ArrayEx0
    Dim rearr As New ArrayEx0
    If IsMissing(stringFormat) Then
        Set ToString = rearr.Init(Format(array_, "@"))
    Else
        Set ToString = rearr.Init(Format(array_, stringFormat))
    End If
End Function

