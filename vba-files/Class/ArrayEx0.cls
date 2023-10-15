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

' -------------------------------------------------------------------------
' Get method : return new instance.
' Set method : return updated current instance.
' Unexpected error, ArrayEx should return instance but null value.
' -------------------------------------------------------------------------

' Fields ------------------------------------------------------------------------------
Private array_ As Variant

' Properties -------------------------------------------------------------------------
Public Property Get Value()
Attribute Value.VB_UserMemId = 0
    Value = array_
End Property


' Public Methods --------------------------------------------------------------------
Public Function Init(array0d) As ArrayEx0
    array_ = array0d
    Set Init = Me
End Function

Public Function Create(Optional arr As Variant) As ArrayEx0
    Dim rearr As New ArrayEx0
    Set Create = IIf(IsMissing(arr), rearr.Init(array_), rearr.Init(arr))
End Function

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

Public Function ToLong() As ArrayEx0
    Dim rearr As New ArrayEx0
    If IsNumeric(array_) Then
        Set ToLong = rearr.Init(CLng(array_))
    Else
        Set ToLong = rearr.Init(Null)
    End If
End Function

Public Function ToDouble() As ArrayEx0
    Dim rearr As New ArrayEx0
    If IsNumeric(array_) Then
        Set ToDouble = rearr.Init(CDbl(array_))
    Else
        Set ToDouble = rearr.Init(Null)
    End If
End Function

Public Function ToDate() As ArrayEx0
    Dim rearr As New ArrayEx0
    If IsDate(array_) Then
        Set ToDate = rearr.Init(CDate(array_))
    Else
        Set ToDate = rearr.Init(Null)
    End If
End Function

