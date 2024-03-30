VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UnitTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private errNumber_ As Long
Private errDescription_ As String
Private errSource_ As String
Private errResult_ As String

Private testName_ As String
Private testNames_ As Collection

Private Property Get isSingleTest() As Boolean
    isSingleTest = True
    If testNames_ Is Nothing Then Exit Property
    If testNames_.Count = 0 Then Exit Property
    isSingleTest = False
End Property

Private Sub Class_Initialize()
    Set testNames_ = New Collection
End Sub

' ----------------------------------------------------------------------------------------------------------------------
' before using, switch error check in Tools > Options > General > Error trap
Public Function RunTests()
    Dim testName
    Debug.Print "--- Start tests (" & Format(Now) & ") ---"
    For Each testName In testNames_
        RunTest (testName)
    Next
    Debug.Print "--- Finish tests (" & Format(Now) & ") ---"
    Set testNames_ = Nothing
End Function

Public Function RegisterTest(testName As String)
    testNames_.Add testName
End Function

Public Function RunTest(testName As String)
On Error Resume Next
    Application.Run (testName)
    CopyErr "RunTest" ' When application run has error, catch
    Debug.Print (IIf(errNumber_ = 0, "OK", "NG") & ": " & testName & IIf(errNumber_ = 0, "", errResult_))
    ClearErr
End Function


Public Function AssertTrue(condition)
    On Error GoTo errCondition
    If condition = True Then Else Call Err.Raise(9001, "", "NG: Should be True.")
    Exit Function
errCondition:
    CopyErr "AssertTrue"
    If isSingleTest Then DebugPrintErr
End Function

Public Function AssertFalse(condition)
    On Error GoTo errCondition
    If condition = False Then Else Call Err.Raise(9002, "", "NG: Should be False.")
    Exit Function
errCondition:
    CopyErr "AssertFalse"
     If isSingleTest Then DebugPrintErr
End Function

Public Function AssertEqual(expected, actual)
    On Error GoTo errCondition
    If VarType(expected) = vbObject And VarType(actual) = vbObject Then If expected.Equals((actual)) Then GoTo pass Else Call Err.Raise(9003, "", "NG: Should be equal.")
    If expected = actual Then Else Call Err.Raise(9003, "", "NG: Should be equal. expected is " & expected & ", actual is " & actual)
pass:
    Exit Function
errCondition:
    CopyErr "AssertEqual"
    If isSingleTest Then DebugPrintErr
End Function

Public Function AssertNotEqual(expected, actual)
    On Error GoTo errCondition
    If expected <> actual Then Else Call Err.Raise(9004, "", "NG: Should not be equal. expected is " & expected & ", actual is " & actual)
    Exit Function
errCondition:
    CopyErr "AssertNotEqual"
    If isSingleTest Then DebugPrintErr
End Function

Public Function AssertHasError()
     If Err.Number <> 0 Then
    Else
        On Error GoTo errCondition
        Call Err.Raise(9005, "", "NG: Should have error, but has no error.")
    End If
    Err.Clear
    Exit Function
errCondition:
    CopyErr "AssertHasError"
    If isSingleTest Then DebugPrintErr: Err.Clear
End Function

Public Function AssertHasNoError()
    If Err.Number = 0 Then
    Else
        Dim num As Long:    num = Err.Number
        Dim desc As String:  desc = Err.Description
        Dim src As String:     src = Err.Source
        On Error GoTo errCondition
        Call Err.Raise(9006, "", "NG: Should have no error, but has error." & vbCrLf & " >>> " & num & ", " & desc & ", " & src)
    End If
    Exit Function
errCondition:
    CopyErr "AssertHasNoError"
    If isSingleTest Then DebugPrintErr:     Err.Clear
End Function

' ------------------------------------------------------------------------
Private Function DebugPrintErr()
    Debug.Print errSource_ & " : " & errDescription_
End Function

Private Function CopyErr(Optional testName As String)
    errNumber_ = IIf(Err.Number = 0, errNumber_, Err.Number)
    errDescription_ = IIf(Err.Number = 0, errDescription_, Err.Description)
    errSource_ = IIf(Err.Number = 0, errSource_, "(" & testName & ") " & Err.Source)
    errResult_ = IIf(Err.Number = 0, errResult_, errResult_ & vbCrLf & "  - " & errSource_ & " : " & errDescription_)
End Function

Private Function ClearErr()
    errNumber_ = 0
    errDescription_ = ""
    errSource_ = ""
    errResult_ = ""
End Function


