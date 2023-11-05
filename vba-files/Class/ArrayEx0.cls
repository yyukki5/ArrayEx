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
Attribute Value.VB_UserMemId = 0
    Value = array_
End Property

Public Property Let Value(val As Variant)
    Init val
End Property

' [Hack] need to more good idea.
' for argument.  need for return value as ArrayEx0
Public Property Get Mine()
    Set Mine = Me
End Property

' Public Methods ---------------------------------------------------------------------------------------------
Public Function Init(array0d As Variant) As ArrayEx0
    If IsArrayRank(array0d, 0) = False Then Call Err.Raise(101, "ArrayEx0 Init", "ArrayEx: Need 0d array (value)")
    array_ = array0d
    Set Init = Me
End Function

Public Function Create(Optional array0d As Variant) As ArrayEx0
    If IsMissing(array0d) = True And IsEmpty(array_) Then Call Err.Raise(105, "ArrayEx0 Create", "ArrayEx: Not initialized")
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

