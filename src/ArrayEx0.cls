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
'    If ArrayExCore.Rank(array0d) = 0 Then
'        array_ = array0d
'    Else
'        Call Err.Raise(101, "ArrayEx0 Init", "ArrayEx: Need 0d array (value)")
'    End If

'    array_ = ArrayExCore.ConvertToDimensionN(array0d, 0)

    If ArrayExCore.Rank(array0d) <> 0 Then Call Err.Raise(101, "ArrayEx0 Init", "ArrayEx: Need 0d array (value)")
    array_ = array0d
    Set Init = Me
End Function

Public Function Create(Optional array0d As Variant) As ArrayEx0
    If IsMissing(array0d) = True And IsEmpty(array_) Then Err.Raise 105, "ArrayEx0 Create", "ArrayEx: Not initialized"
    Dim rearr As New ArrayEx0
    If IsMissing(array0d) Then
        Set Create = rearr.Init(array_)
    Else
        Set Create = rearr.Init(array0d)
    End If
End Function

Public Function Equals(array0d) As Boolean
    Equals = (array_ = array0d)
End Function

' Convert
Public Function ToString(Optional stringFormat As String) As String
    ToString = CStr(array_)
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

