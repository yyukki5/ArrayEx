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

' ------------------------------------------------------------------------------------------------------------
' Init methods              : should be initialize by array.
' Convert methods           : should return other type object.
' DebugPrint                : should return current instance. print current array value in immediate window.
' Wrap default functions    : should return new instance.
' ------------------------------------------------------------------------------------------------------------

' Fields -----------------------------------------------------------------------------------------------------
Private array_ As Variant

' Properties -------------------------------------------------------------------------------------------------
Public Property Get Value()
    Value = array_
End Property

Public Property Let Value(val As Variant)
    Init val
End Property

' Public Methods ---------------------------------------------------------------------------------------------
Public Function Init(array0d As Variant) As ArrayEx0
    ArrayExCore.Validate array0d
    If ArrayExCore.Rank(array0d) <> 0 Then Call Err.Raise(1001, "ArrayEx0.Init", "ArrayEx: Need 0d array (value)")
    array_ = array0d
    Set Init = Me
End Function

Public Function InitForce(arr As Variant) As ArrayEx0
    array_ = ArrayExCore.ConvertToDimensionN(arr, 0)
    Set InitForce = Me
End Function

'' todo delete ?
'Public Function Create(Optional array0d As Variant) As ArrayEx0
'    If IsMissing(array0d) = True And IsEmpty(array_) Then Err.Raise 105, "ArrayEx0 Create", "ArrayEx: Not initialized"
'    Dim rearr As New ArrayEx0
'    If IsMissing(array0d) Then
'        Set Create = rearr.Init(array_)
'    Else
'        Set Create = rearr.Init(array0d)
'    End If
'End Function

Public Function Equals(arrayEx0d As ArrayEx0) As Boolean
    Equals = (array_ = arrayEx0d.Value)
End Function

' Convert
Public Function ToString(Optional formatText As String = "", Optional FirstDayOfWeek As Long = vbSunday, Optional FirstWeekOfYear As Long = vbFirstJan1) As String
    ToString = Format(CStr(array_), formatText, FirstDayOfWeek, FirstWeekOfYear)
End Function

' Debug
Public Function DebugPrint() As ArrayEx0
    Debug.Print array_
    Set DebugPrint = Me
End Function

' wrap default functions
Public Function Replace(before As String, after As String) As ArrayEx0
    Dim rearr As New ArrayEx0
    Set Replace = rearr.Init(Strings.Replace(array_, before, after))
End Function

Public Function Left(length As Long) As ArrayEx0
    Dim rearr As New ArrayEx0
    Set Left = rearr.Init(Strings.Left(array_, length))
End Function

Public Function Right(length As Long) As ArrayEx0
    Dim rearr As New ArrayEx0
    Set Right = rearr.Init(Strings.Right(array_, length))
End Function

Public Function Mid(start As Long, Optional length As Long = 0) As ArrayEx0
    Dim rearr As New ArrayEx0
    If length = 0 Then
        Set Mid = rearr.Init(Strings.Mid(array_, start))
    Else
        Set Mid = rearr.Init(Strings.Mid(array_, start, length))
    End If
End Function

' [Hack] if need more, add


' Private Methods --------------------------------------------------------------------------------------------

