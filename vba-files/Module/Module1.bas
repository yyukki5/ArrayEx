Attribute VB_Name = "Module1"
'Sub test1()
'
'    Dim a As New ArrayEx2
'    Set cols = a.Init(Range("A1:E5").Value).Extract("2 to 4").ToCollection()
'
'
'    Dim ccc2 As New ArrayEx2
'    Set ccc2 = ccc2.CreateFromCollection(cols)
'
'
'
'    Dim col As ArrayEx1
'    For Each col In cols
'        Debug.Print "ID: " & col(1) & " Hello World" & col(2)
'        col.GetItem(1).ToString
'
'        Dim c As ArrayEx0
'        For Each c In col.ToCollection
'
'            Debug.Print c.ToString().Replaced("1", "5").ToString("""ID"" 00.00")
'        Next c
'
'    Next col
'
'
'End Sub

Sub test2()

Dim abc As New ArrayEx2

On Error GoTo myerr
abc.Init (Array(1, 2, 3))

myerr:
    Debug.Print Err.Description
End Sub


Sub test3()

Dim abc As New ArrayEx1, bcd As ArrayEx1
abc.Init (Array(1, 2, 3))

With abc
    Debug.Print .Value(1)
    .GetItem (1)
    .ToCollection

End With


End Sub


'Sub test4()
'    Dim a As New ArrayEx2
'    Dim b
'    Set b = a.Init(Range("A1:E5").Value) _
'    .GetRows(1) _
'    .DebugPrintEvaluated(" ""result of {x}^2 :"" & {x}^2 & "" ---"" ") _
'    .ToCollection
'
'    Dim e As New ArrayEx1
'
'    Set e = a.Init(Range("A1:E5").Value) _
'    .GetRows(1) _
'    .DebugPrintEvaluated("{x}^2") _
'    .WhereEvaluated("mod({x},3) = 0 ") _
'    .DebugPrintEvaluated(" ""_"" & {x}^3  ") _
'    .SelectEvaluated("TEXT({x} , ""aaa"") ") _
'    .DebugPrint("___")
'
''    For i = e.Lb To e.Ub
''        Debug.Print e(i)
''    Next i
'
'
'
'
'    Dim c As New ArrayEx0
'    Dim d As New Collection
'    For Each c In b
'        d.Add c.Replaced("1", "2").Value
'    Next c
'
''    Dim e As New ArrayEx1
''    e.CreateFromCollection (d)....
'
'End Sub

Sub testabc()

s = "3 : 12"
s = LCase(Replace(s, " ", ""))
If InStr(s, "to") > 0 Then
    l = val(Left(s, InStr(s, "to") - 1))
    u = val(Right(s, InStrRev(s, "to")))
End If
If InStr(s, ":") > 0 Then
    l = val(Left(s, InStr(s, ":") - 1))
    u = val(Right(s, InStrRev(s, ":")))
End If


Dim re
ReDim re(1 To u - l + 1)
Dim i As Long, j As Long: j = 1
For i = l To u
    re(j) = i
    j = j + 1
Next i
End Sub

'
'
'Sub aaaaa()
'    Dim a As New ArrayEx2
'    Set a = a.Init(Range("J1:L5").Value) _
'        .CreateHeader() _
'        .DebugPrintEvaluated("{x},{y}", "1,2", " ""Debug Print --- ID: "" & {x} & "" Name: {y}"" ")
'
'    Dim abc As Variant
'
'     abc = [2 to 4]
'
'    Set col = a.Extract("2 to 4") _
'        .ToCollection()
'
'    Dim b As ArrayEx1
'    For Each b In col
'        Debug.Print "ID:" & b(a.GetCol("id")) & " Name: " & b(a.GetCol("comment"))
'    Next b
'
'End Sub


Sub test5()

    Dim a As New ArrayEx2
    Dim rows
    Dim arr
    arr = Array(2, 3)
    Set rows = a.Init(Range("A1:E5").Value) _
                          .DebugPrint("{x}, {y}", "1,2", "ID:{x} - {y}") _
                          .GetRows(Array(2, 5)) _
                          .DebugPrint("{x},{y},z", "1,2,3", "ID:{x} - {y} - z") _
                          .GetColumns(Array(1, 2)) _
                          .DebugPrint("x,y", "1,2", "1:x, 2:y")
                    
End Sub


Sub test6()
    Dim starttime
    starttime = Timer
    Dim a As New ArrayEx2
    Dim rows
    Dim arr: arr = Range("A1:E5").Value
    Dim i
    For i = 1 To 1000
        Set rows = a.Init(arr) _
                              .GetRows(Array(2, 5)) _
                              .GetColumns(Array(1, 2))
    Next i
    Debug.Print (Timer - starttime) * 1000 & "[ms]"
End Sub

Sub test7()
    Dim starttime
    starttime = Timer
    Dim a As New ArrayEx2
    Dim rows
    Dim arr: arr = Range("A1:E5").Value
    Dim i
    For i = 1 To 100
        Set rows = a.Init(arr).Extract(Array(2, 3, 4, 5), Array(1, 2))
'        Set rows = a.Init(arr).Extract("2 to 5", "1 to 2")
'        Set rows = a.Init(arr).Extract("2:5", ":").GetColumn(2).DebugPrintAllItems.GetItem(4).ToString
    Next i
    Debug.Print (Timer - starttime) * 1000 & "[ms]"
End Sub


Sub test8()

    Dim a As New ArrayEx2
    Dim rows, rows1
    Dim arr
    arr = Array(2, 3)
    Set rows = a.Init(Range("A1:E5").Value).DebugPrintAll
    Call rows.SetValue(":", 2, 5)
    Call rows.DebugPrintAll
    Call rows.SetValue(4, ":", 5)
    Call rows.DebugPrintAll
    Call rows.SetValue(":", ":", 0)
    Call rows.DebugPrintAll
End Sub

Sub test9()

    Dim a As New ArrayEx2
    Dim rows As New ArrayEx2, rows1
    Dim arr
    arr = Array(2, 3)
    Set rows = a.Init(Range("A1:E5").Value).DebugPrintAll
    
    Call rows.SetValue("3,4", ":", 5)
    
    rows.DebugPrintAll _
        .GetRows("1,2:3").DebugPrintAll _
        .GetColumns("2,3").DebugPrintAll
    
End Sub

Sub test10()

    Dim a As New ArrayEx2
    Dim rows As New ArrayEx2, rows1
    Set rows = a.Init(Range("A1:E5").Value).DebugPrintAll
    Debug.Print a.GetElement(3, 1).ToString
    Dim b, c, d, e
    b = rows()
    
    b = rows(1, 3)
    c = rows("1to")
    d = rows(, "2")
    e = rows("1:2", "1 to 2")
    
End Sub

Sub test11()

    Dim a As New ArrayEx2
    Dim rows As New ArrayEx2, rows1
    Set rows = a.Init(Range("A1:E5").Value).DebugPrintAll
    
    Call rows.SetRow(1, rows.GetRow(3))
    rows.DebugPrintAll
    Call rows.SetColumn(3, rows.GetColumn(1))
    rows.DebugPrintAll
    
    Debug.Assert rows.GetColumn(1).Max = 5
    Debug.Assert rows.GetColumn(1).Min = 2

    
End Sub

Sub test12()
    Dim a As New ArrayEx2
    Dim b
    Call a.Init(Range("A1:E5").Value).DebugPrintAll _
        .GetColumns("3,2,1").DebugPrintAll _
        .Transpose.DebugPrintAll _
        .SetValue(1, 1, 5) _
        .ToRange(Range("H10"))
End Sub


Public Function SearchWord(ary, word As String) As Variant
    Dim a As New ArrayEx2: a.Init (ary)
    Dim i As Long, s As String
    For i = a.Lb To a.Ub
        If InStr(WorksheetFunction.TextJoin(",", True, a.GetRow(i).Value), word) > 0 Then
        s = s & i & ","
        End If
    Next i
    If s = "" Then Exit Function
    s = Left(s, Len(s) - 1)
    SearchWord = a.GetRows(s).Value
End Function
