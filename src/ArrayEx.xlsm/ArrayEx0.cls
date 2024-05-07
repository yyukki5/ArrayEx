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
' Project                   : https://github.com/yyukki5/ArrayEx
' Version                   : 2.2.0
' Dependencies              : ArrayExCore.cls
' ------------------------------------------------------------------------------------------------------------

' Fields -----------------------------------------------------------------------------------------------------
Private array_ As Variant


' Properties -------------------------------------------------------------------------------------------------
Public Property Get Value(): Value = array_: End Property
Public Property Let Value(val As Variant): Init val: End Property

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

Public Function Equals(arrayEx0d As ArrayEx0) As Boolean
    Equals = (array_ = arrayEx0d.Value)
End Function

' Convert
Public Function ToString(Optional formatText, Optional FirstDayOfWeek As VbDayOfWeek = vbSunday, Optional FirstWeekOfYear As VbFirstWeekOfYear = vbFirstJan1) As String
    ToString = Strings.Format(CStr(array_), IIf(IsMissing(formatText), "", formatText), FirstDayOfWeek, FirstWeekOfYear)
End Function
Public Function ConvertTo(t As VbVarType) As ArrayEx0
    Dim rearr As New ArrayEx0, re
    Select Case t
        Case VbVarType.vbBoolean:    Set ConvertTo = rearr.Init(CBool(array_))
        Case VbVarType.vbByte:       Set ConvertTo = rearr.Init(CByte(array_))
        Case VbVarType.vbDouble:     Set ConvertTo = rearr.Init(CDbl(array_))
        Case VbVarType.vbCurrency:   Set ConvertTo = rearr.Init(CCur(array_))
        Case VbVarType.vbDate:      Set ConvertTo = rearr.Init(CDate(array_))
        Case VbVarType.vbDecimal:   Set ConvertTo = rearr.Init(CDec(array_))
        Case VbVarType.vbInteger:   Set ConvertTo = rearr.Init(CInt(array_))
        Case VbVarType.vbLong:      Set ConvertTo = rearr.Init(CLng(array_))
        Case VbVarType.vbSingle:    Set ConvertTo = rearr.Init(CSng(array_))
        Case VbVarType.vbString:    Set ConvertTo = rearr.Init(CStr(array_))
        Case Else:                  Set ConvertTo = rearr.Init(array_)
    End Select
End Function

' Debug
Public Function DebugPrint() As ArrayEx0
    Debug.Print array_
    Set DebugPrint = Me
End Function

' wrap default functions

' Strings
Public Function Format(Optional formatText, Optional FirstDayOfWeek As VbDayOfWeek = vbSunday, Optional FirstWeekOfYear As VbFirstWeekOfYear = vbFirstJan1) As ArrayEx0
    Dim rearr As New ArrayEx0:     Set Format = rearr.Init(Strings.Format(CStr(array_), IIf(IsMissing(formatText), "", formatText), FirstDayOfWeek, FirstWeekOfYear))
End Function
Public Function Replace(before As String, after As String) As ArrayEx0
    Dim rearr As New ArrayEx0:    Set Replace = rearr.Init(Strings.Replace(array_, before, after))
End Function
Public Function Left(length As Long) As ArrayEx0
    Dim rearr As New ArrayEx0:    Set Left = rearr.Init(Strings.Left(array_, length))
End Function
Public Function Right(length As Long) As ArrayEx0
    Dim rearr As New ArrayEx0:    Set Right = rearr.Init(Strings.Right(array_, length))
End Function
Public Function Mid(start As Long, Optional length) As ArrayEx0
    Dim rearr As New ArrayEx0:    Set Mid = rearr.Init(Strings.Mid(array_, start, IIf(IsMissing(length), Strings.Len(array_), length)))
End Function
Public Function Concat(text As String) As ArrayEx0
    Dim rearr As New ArrayEx0:    Set Concat = rearr.Init(array_ & text)
End Function
Public Function InStr(Optional start, Optional string2, Optional compare As VbCompareMethod = VbCompareMethod.vbBinaryCompare) As ArrayEx0
    Dim rearr As New ArrayEx0:    Set InStr = rearr.Init(Strings.InStr(IIf(IsNumeric(start), start, 1), array_, IIf(IsNumeric(start), string2, start), compare))
End Function
Public Function InStrRev(stringmatch, Optional start As Long = -1, Optional compare As VbCompareMethod = vbBinaryCompare) As ArrayEx0
    Dim rearr As New ArrayEx0:    Set InStrRev = rearr.Init(Strings.InStrRev(array_, stringmatch, start, compare))
End Function
Public Function StrComp(str, Optional compare As VbCompareMethod = vbBinaryCompare) As ArrayEx0
    Dim rearr As New ArrayEx0:    Set StrComp = rearr.Init(Strings.StrComp(array_, str, compare))
End Function
Public Function StrConv(Conversion As VbStrConv) As ArrayEx0
    Dim rearr As New ArrayEx0:    Set StrConv = rearr.Init(Strings.StrConv(array_, Conversion))
End Function
Public Function StrReverse() As ArrayEx0
    Dim rearr As New ArrayEx0:    Set StrReverse = rearr.Init(Strings.StrReverse(CStr(array_)))
End Function
Public Function Asc() As ArrayEx0
    Dim rearr As New ArrayEx0:    Set Asc = rearr.Init(Strings.Asc(CStr(array_)))
End Function
Public Function Chr() As ArrayEx0
    Dim rearr As New ArrayEx0:    Set Chr = rearr.Init(Strings.Chr(CLng(array_)))
End Function
Public Function Trim() As ArrayEx0
    Dim rearr As New ArrayEx0:    Set Trim = rearr.Init(Strings.Trim(array_))
End Function
Public Function LTrim() As ArrayEx0
    Dim rearr As New ArrayEx0:    Set LTrim = rearr.Init(Strings.LTrim(array_)):
End Function
Public Function RTrim() As ArrayEx0
    Dim rearr As New ArrayEx0:    Set RTrim = rearr.Init(Strings.RTrim(array_)):
End Function


Public Function Add(a) As ArrayEx0
    Dim rearr As New ArrayEx0
    Set Add = rearr.Init(array_ + a)
End Function
Public Function Multiple(a) As ArrayEx0
    Dim rearr As New ArrayEx0
    If IsNumeric(a) = False Then Err.Raise 9100, "ArrayEx0.Multiple()", "value is not numeric"
    Set Multiple = rearr.Init(array_ * a)
End Function

' [Hack] Wrap default functions. Better to use app.*** in case of depending on application function.


' Private Methods --------------------------------------------------------------------------------------------
