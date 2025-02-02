Attribute VB_Name = "UnitTests"
'<dir .\Tests /dir>
' This module depended on UnitTest.cls
' Should be passed all test of RunTests().

Option Explicit
Option Private Module

Private test_ As New UnitTest

Sub Create_Tests()
    Call UnitTest.CreateRunTests
End Sub

Sub RunTests()
   Dim test As New UnitTest

    test.RegisterTest "Initialize__"
    test.RegisterTest "Initialize_Null_"
    test.RegisterTest "Equals_"
    test.RegisterTest "RedimPreserve_0toN_"
    test.RegisterTest "RedimPreserve_1toN_"
    test.RegisterTest "RedimPreserve_2toN_"
    test.RegisterTest "VerticalStack_2DimOnly"
    test.RegisterTest "VerticalStack_MixDim"
    test.RegisterTest "HorizontalStack_2DimOnly"
    test.RegisterTest "HorizontalStack_MixDim"
    test.RegisterTest "ShiftIndex_2Dim"
    test.RegisterTest "SelectColumns_2Dim_"
    test.RegisterTest "SelectColumns_1Dim_"
    test.RegisterTest "SelectRows_2Dim_"
    test.RegisterTest "SelectColumn_2Dim_"
    test.RegisterTest "SelectColumn_01Dim_"
    test.RegisterTest "SelectRow_2Dim_"
    test.RegisterTest "SelectRow_01Dim_"
    test.RegisterTest "Where_2Dim_"
    test.RegisterTest "Where_1Dim_"
    test.RegisterTest "FindIndex_2Dim_"
    test.RegisterTest "FindIndex_1Dim_"
    test.RegisterTest "Skip_1Dim_"
    test.RegisterTest "Skip_2Dim_"
    test.RegisterTest "Take_2Dim_"
    test.RegisterTest "OrderBy_1Dim_Number"
    test.RegisterTest "OrderBy_1Dim_String"
    test.RegisterTest "OrderByDescending_1Dim_Number"
    test.RegisterTest "OrderByDescending_1Dim_String"
    test.RegisterTest "OrderBy_2Dim"
    test.RegisterTest "OrderByDescending_2Dim"
    test.RegisterTest "Distinct_0Dim_"
    test.RegisterTest "Distinct_1Dim_"
    test.RegisterTest "Distinct_2Dim_"
    test.RegisterTest "DistinctBy_2Dim_"
    test.RegisterTest "Transpose__"
    test.RegisterTest "XLookUp_1Dim_"
    test.RegisterTest "XLookUp_2Dim_"
    test.RegisterTest "Cast__"
    test.RegisterTest "ToCollection_1Dim_"
    test.RegisterTest "ToCollection_2Dim_"
    test.RegisterTest "InnerJoin__"
    test.RegisterTest "LeftJoin__"
    test.RegisterTest "FullOuterJoin__"
    test.RegisterTest "CrossJoin__"

    test.RunTests UnitTest
End Sub


'[Fact]
Sub Initialize__()
        
    ' Arrange
    Dim a1, b, aex As New ArrEx
    a1 = array1d_1to10
    b = a1(2)
    
    ' Act
    Call aex.Initialize(a1)
    
    ' Assert
    With UnitTest.NameOf("Initailzie")
        Call .AssertEqual(a1(1), aex.Value(1))
        
        a1(2) = 100
        Call .AssertEqual(b, aex.Value(2))
    End With

End Sub

'[Fact]
Sub Initialize_Null_()
        
    ' Arrange
    Dim a1, aex As New ArrEx
    a1 = Null
    
    ' Act
    Call aex.Initialize(a1)
    
    ' Assert
    With UnitTest.NameOf("Initailzie_Null")
        Call .AssertEqual(Null, aex.Value)
        Call .AssertEqual(0, aex.Count)
    End With

End Sub

'[Fact]
Sub Equals_()
        
    ' Arrange
    Dim a1, aex As New ArrEx
    
    ' Act & Assert
    With UnitTest.NameOf("Equals")
        Call .AssertTrue(aex.Create(Null).Equals(Null))
        Call .AssertTrue(aex.Create(1).Equals(1))
        Call .AssertFalse(aex.Create(1).Equals(2))
        Call .AssertTrue(aex.Create("abc").Equals("abc"))
        Call .AssertTrue(aex.Create(array1d_1to10).Equals(array1d_1to10))
        Call .AssertFalse(aex.Create(array1d_1to10).Equals(array1d_1to10_UnSort))
        Call .AssertTrue(aex.Create(array2d_1to3x0to12).Equals(array2d_1to3x0to12))
        
        Call .AssertTrue(aex.Create(array2d_1to3x0to12).Equals(aex.Create(array2d_1to3x0to12)))
    End With

End Sub

'[Fact]
Sub RedimPreserve_0toN_()

    ' Arrange
    Dim a, arr1, arr2
    a = 1
        
    ' Act & Assert
    With UnitTest.NameOf("Redim 0 to N")
        Call .AssertEqual(0, ArrEx(a).RedimPreserve(aexRank0).Rank)
        Call .AssertEqual(a, ArrEx(a).RedimPreserve(aexRank0).Value)
        
        Call .AssertEqual(1, ArrEx(a).RedimPreserve(aexRank1).Rank)
        Call .AssertEqual(a, ArrEx(a).RedimPreserve(aexRank1).Value(0))
        Call .AssertEqual(1, ArrEx(a).RedimPreserve(aexRank1, 0, 1).Rank)
        Call .AssertEqual(1, ArrEx(a).RedimPreserve(aexRank1, 1, 2).LowerBound)
        Call .AssertEqual(5, ArrEx(a).RedimPreserve(aexRank1, 1, 5).UpperBound)
        Call .AssertEqual(9, ArrEx(a).RedimPreserve(aexRank1, 1, 5, blank_value:=9).Value(2))
        
        Call .AssertEqual(2, ArrEx(a).RedimPreserve(aexRank2).Rank)
        Call .AssertEqual(a, ArrEx(a).RedimPreserve(aexRank2).Value(0, 0))
        Call .AssertEqual(a, ArrEx(a).RedimPreserve(aexRank2, 0, 1).Value(0, 0))
        Call .AssertEqual(1, ArrEx(a).RedimPreserve(aexRank2, 1, 2).LowerBound)
        Call .AssertEqual(1, ArrEx(a).RedimPreserve(aexRank2, 0, 1).UpperBound)
        Call .AssertEqual(0, ArrEx(a).RedimPreserve(aexRank2, 0, 1).UpperBound(2))
        Call .AssertEqual(2, ArrEx(a).RedimPreserve(aexRank2, 0, 1, 0, 1).Rank)
        Call .AssertEqual(a, ArrEx(a).RedimPreserve(aexRank2, 0, 1, 0, 1).Value(0, 0))
        Call .AssertEqual(1, ArrEx(a).RedimPreserve(aexRank2, 1, 2, 3, 5).LowerBound(1))
        Call .AssertEqual(2, ArrEx(a).RedimPreserve(aexRank2, 1, 2, 3, 5).UpperBound(1))
        Call .AssertEqual(3, ArrEx(a).RedimPreserve(aexRank2, 1, 2, 3, 5).LowerBound(2))
        Call .AssertEqual(5, ArrEx(a).RedimPreserve(aexRank2, 1, 2, 3, 5).UpperBound(2))
        Call .AssertEqual(9, ArrEx(a).RedimPreserve(aexRank2, 1, 2, 3, 5, 9).Value(2, 5))
        
        On Error Resume Next
        Call ArrEx(a).RedimPreserve(aexRank0, 2)
        Call .AssertHasError
        Call ArrEx(a).RedimPreserve(aexRank1, 2)
        Call .AssertHasError
        Call ArrEx(a).RedimPreserve(aexRank2, 1, 2, 3)
        Call .AssertHasError

    End With
        
End Sub

'[Fact]
Sub RedimPreserve_1toN_()

    ' Arrange
    Dim a
    a = array1d_1to10
        
    ' Act & Assert
    With UnitTest.NameOf("Redim 1 to N")
        Call .AssertEqual(0, ArrEx(a).RedimPreserve(aexRank0).Rank)
        Call .AssertEqual(a(1), ArrEx(a).RedimPreserve(aexRank0).Value)
        
        Call .AssertEqual(1, ArrEx(a).RedimPreserve(aexRank1, 0, 10).Rank)
        Call .AssertEqual(0, ArrEx(a).RedimPreserve(aexRank1, 0, 10).LowerBound)
        Call .AssertEqual(10, ArrEx(a).RedimPreserve(aexRank1, 0, 10).UpperBound)
        Call .AssertEqual(a(1), ArrEx(a).RedimPreserve(aexRank1, 0, 10).Value(0))
        Call .AssertEqual(100, ArrEx(a).RedimPreserve(aexRank1, 0, 10, blank_value:=100).Value(10))
        Call .AssertEqual(2, ArrEx(a).RedimPreserve(aexRank1, 2, 5).LowerBound)
        Call .AssertEqual(5, ArrEx(a).RedimPreserve(aexRank1, 2, 5).UpperBound)
        Call .AssertEqual(a(1), ArrEx(a).RedimPreserve(aexRank1, 2, 5).Value(2))
        
        Call .AssertEqual(2, ArrEx(a).RedimPreserve(2, 0, 10, 0, 10).Rank)
        Call .AssertEqual(0, ArrEx(a).RedimPreserve(2, 0, 10, 0, 10).LowerBound(1))
        Call .AssertEqual(10, ArrEx(a).RedimPreserve(2, 0, 10, 0, 10).UpperBound(1))
        Call .AssertEqual(0, ArrEx(a).RedimPreserve(2, 0, 10, 0, 10).LowerBound(2))
        Call .AssertEqual(10, ArrEx(a).RedimPreserve(2, 0, 10, 0, 10).UpperBound(2))
        Call .AssertEqual(a(1), ArrEx(a).RedimPreserve(aexRank2, 0, 10, 0, 10).Value(0, 0))
        Call .AssertEqual(0, ArrEx(a).RedimPreserve(aexRank2, 0, 10, 0, 10, 0).Value(10, 10))
        Call .AssertEqual(1, ArrEx(a).RedimPreserve(2, 1, 2, 3, 5).LowerBound(1))
        Call .AssertEqual(2, ArrEx(a).RedimPreserve(2, 1, 2, 3, 5).UpperBound(1))
        Call .AssertEqual(3, ArrEx(a).RedimPreserve(2, 1, 2, 3, 5).LowerBound(2))
        Call .AssertEqual(5, ArrEx(a).RedimPreserve(2, 1, 2, 3, 5).UpperBound(2))
        Call .AssertEqual(a(1), ArrEx(a).RedimPreserve(aexRank2, 1, 1, 3, 5).Value(1, 3))
        Call .AssertEqual(a(3), ArrEx(a).RedimPreserve(aexRank2, 1, 1, 3, 5).Value(1, 5))
                        
    End With

End Sub

'[Fact]
Sub RedimPreserve_2toN_()

    ' Arrange
    Dim a
    a = array2d_1to3x0to12
        
    ' Act & Assert
    With UnitTest.NameOf("Redim 2 to N")
        Call .AssertEqual(0, ArrEx(a).RedimPreserve(aexRank0).Rank)
        Call .AssertEqual(a(1, 0), ArrEx(a).RedimPreserve(aexRank0).Value)
        
        Call .AssertEqual(1, ArrEx(a).RedimPreserve(aexRank1, 0, 10).Rank)
        Call .AssertEqual(0, ArrEx(a).RedimPreserve(aexRank1, 0, 15).LowerBound)
        Call .AssertEqual(15, ArrEx(a).RedimPreserve(aexRank1, 0, 15).UpperBound)
        Call .AssertEqual(a(1, 0), ArrEx(a).RedimPreserve(aexRank1, 0, 15).Value(0))
        Call .AssertEqual(0, ArrEx(a).RedimPreserve(aexRank1, 0, 15, blank_value:=0).Value(15))
        Call .AssertEqual(2, ArrEx(a).RedimPreserve(aexRank1, 2, 5).LowerBound)
        Call .AssertEqual(5, ArrEx(a).RedimPreserve(aexRank1, 2, 5).UpperBound)
        Call .AssertEqual(a(1, 0), ArrEx(a).RedimPreserve(aexRank1, 2, 5).Value(2))
        Call .AssertEqual(a(1, 3), ArrEx(a).RedimPreserve(aexRank1, 2, 5).Value(5))
        Call .AssertEqual(1, ArrEx(a).RedimPreserve(aexRank1).Rank)
        Call .AssertEqual(0, ArrEx(a).RedimPreserve(aexRank1).LowerBound)
        Call .AssertEqual(12, ArrEx(a).RedimPreserve(aexRank1).UpperBound)
        Call .AssertEqual(a(1, 12), ArrEx(a).RedimPreserve(aexRank1).Value(12))
        
        Call .AssertEqual(2, ArrEx(a).RedimPreserve(2, 0, 10, 0, 15).Rank)
        Call .AssertEqual(0, ArrEx(a).RedimPreserve(2, 0, 10, 0, 15).LowerBound(1))
        Call .AssertEqual(10, ArrEx(a).RedimPreserve(2, 0, 10, 0, 15).UpperBound(1))
        Call .AssertEqual(0, ArrEx(a).RedimPreserve(2, 0, 10, 0, 15).LowerBound(2))
        Call .AssertEqual(15, ArrEx(a).RedimPreserve(2, 0, 10, 0, 15).UpperBound(2))
        Call .AssertEqual(a(1, 0), ArrEx(a).RedimPreserve(aexRank2, 0, 10, 0, 15).Value(0, 0))
        Call .AssertEqual(a(3, 12), ArrEx(a).RedimPreserve(aexRank2, 0, 10, 0, 15).Value(2, 12))
        Call .AssertEqual(0, ArrEx(a).RedimPreserve(aexRank2, 0, 10, 0, 15, 0).Value(10, 15))
        Call .AssertEqual(1, ArrEx(a).RedimPreserve(2, 1, 2, 3, 5).LowerBound(1))
        Call .AssertEqual(2, ArrEx(a).RedimPreserve(2, 1, 2, 3, 5).UpperBound(1))
        Call .AssertEqual(3, ArrEx(a).RedimPreserve(2, 1, 2, 3, 5).LowerBound(2))
        Call .AssertEqual(5, ArrEx(a).RedimPreserve(2, 1, 2, 3, 5).UpperBound(2))
        Call .AssertEqual(a(1, 0), ArrEx(a).RedimPreserve(aexRank2, 1, 1, 3, 5).Value(1, 3))
        Call .AssertEqual(a(1, 2), ArrEx(a).RedimPreserve(aexRank2, 1, 1, 3, 5).Value(1, 5))
        
        Call .AssertEqual(a(1, 4), ArrEx(a).RedimPreserve(aexRank2, , , 1, 5).Value(1, 5))
    End With

End Sub


'[Fact]
Sub VerticalStack_2DimOnly()
    
    ' Arrange
    Dim a1, a2, a3, aex As New ArrEx
    a1 = array2d_1to3x0to12
    a2 = array2d_1to3x0to12
    a3 = array2d_1to3x0to12
    Call aex.Initialize(a1)
    
    ' Act
    Set aex = aex.VerticalStack(a2).VerticalStack(a3)
    
    ' Assert
    With UnitTest.NameOf("VerticalStack_2DimOnly")
        Call .AssertEqual(1, aex.LowerBound(1))
        Call .AssertEqual(9, aex.UpperBound(1))
        Call .AssertEqual(0, aex.LowerBound(2))
        Call .AssertEqual(12, aex.UpperBound(2))
        Call .AssertEqual(a1(1, 0), aex.Value(1, 0))
        Call .AssertEqual(a1(3, 12), aex.Value(9, 12))
    End With

End Sub

'[Fact]
Sub VerticalStack_MixDim()
    
    ' Arrange
    Dim a1, a2, a3, aex As New ArrEx
    a2 = array2d_1to3x0to12
    a3 = array1d_1to10
    a1 = 1
    Call aex.Initialize(a1)
    
    ' Act
    Set aex = aex.VerticalStack(a2).VerticalStack(a3, 1)
    
    ' Assert
    With UnitTest.NameOf("VerticalStack_MixDim")
        Call .AssertEqual(0, aex.LowerBound(1))
        Call .AssertEqual(4, aex.UpperBound(1))
        Call .AssertEqual(0, aex.LowerBound(2))
        Call .AssertEqual(12, aex.UpperBound(2))
        Call .AssertEqual(a1, aex.Value(0, 0))
        Call .AssertEqual(Null, aex.Value(0, 2))
        Call .AssertEqual(1, aex.Value(4, 12))
    End With

End Sub

'[Fact]
Sub HorizontalStack_2DimOnly()
    
    ' Arrange
    Dim a1, a2, a3, aex As New ArrEx
    a1 = array2d_1to3x0to12
    a2 = array2d_1to3x0to12
    a3 = array2d_1to3x0to12
    Call aex.Initialize(a1)
    
    ' Act
    Set aex = aex.HorizontalStack(a2).HorizontalStack(a3)
    
    ' Assert
    With UnitTest.NameOf("HorizontalStack_2DimOnly")
        Call .AssertEqual(1, aex.LowerBound(1))
        Call .AssertEqual(3, aex.UpperBound(1))
        Call .AssertEqual(0, aex.LowerBound(2))
        Call .AssertEqual(38, aex.UpperBound(2))
        Call .AssertEqual(a1(1, 0), aex.Value(1, 0))
        Call .AssertEqual(a1(3, 12), aex.Value(3, 38))
    End With

End Sub

'[Fact]
Sub HorizontalStack_MixDim()
    
    ' Arrange
    Dim a1, a2, a3, aex As New ArrEx
    a2 = array2d_1to3x0to12
    a3 = array1d_1to10
    a1 = 1
    Call aex.Initialize(a1)
    
    ' Act
    Set aex = aex.HorizontalStack(a2).HorizontalStack(a3, 1)

    ' Assert
    With UnitTest.NameOf("HorizontalStack_MixDim")
        Call .AssertEqual(0, aex.LowerBound(1))
        Call .AssertEqual(3, aex.UpperBound(1))
        Call .AssertEqual(0, aex.LowerBound(2))
        Call .AssertEqual(23, aex.UpperBound(2))
        Call .AssertEqual(a1, aex.Value(0, 0))
        Call .AssertEqual(Null, aex.Value(2, 0))
        Call .AssertEqual(1, aex.Value(3, 23))
    End With

End Sub

'[Fact]
Sub ShiftIndex_2Dim()
    
    ' Arrange
    Dim a0, a1, a2, aex As New ArrEx
    a0 = 0
    a1 = array1d_1to10
    a2 = array2d_1to3x0to12
    
    ' Act
    Set aex = aex(a2).ShiftIndex(-1, 1)
    
    ' Assert
    With UnitTest.NameOf("ShiftIndex_2Dim")
        Call .AssertEqual(0, aex.LowerBound(1))
        Call .AssertEqual(2, aex.UpperBound(1))
        Call .AssertEqual(1, aex.LowerBound(2))
        Call .AssertEqual(13, aex.UpperBound(2))
    End With
End Sub


'[Fact]
Sub SelectColumns_2Dim_()

    ' Arrange
    Dim a2, aex As New ArrEx
    a2 = array2d_1to3x0to12
    
    ' Act
    Set aex = aex.Initialize(a2).SelectColumns(1, 2, 3)
    
    ' Assert
    With UnitTest.NameOf("SelectColumns_2Dim_")
        Call .AssertEqual(a2(1, 1), aex.Value(1, 0))
        Call .AssertEqual(a2(3, 3), aex.Value(3, 2))
    End With
    
End Sub

'[Fact]
Sub SelectColumns_1Dim_()

    ' Arrange
    Dim a1, aex As New ArrEx
    a1 = array1d_1to10
    
    ' Act
    Set aex = aex.Initialize(a1).SelectColumns(2, 3, 4, 5)
    
    ' Assert
    With UnitTest.NameOf("SelectColumns_1Dim_")
        Call .AssertEqual(a1(2), aex.Value(1))
        Call .AssertEqual(a1(3), aex.Value(2))
        Call .AssertEqual(a1(4), aex.Value(3))
        Call .AssertEqual(a1(5), aex.Value(4))
    End With
    
End Sub


'[Fact]
Sub SelectRows_2Dim_()

    ' Arrange
    Dim a2, aex As New ArrEx
    a2 = array2d_1to3x0to12
    
    ' Act
    Set aex = aex.Initialize(a2).SelectRows(2, 3)

    ' Assert
    With UnitTest.NameOf("SelectRows_2Dim_")
        Call .AssertEqual(a2(2, 0), aex.Value(1, 0))
        Call .AssertEqual(a2(3, 3), aex.Value(2, 3))
    End With
    
End Sub

'[Fact]
Sub SelectColumn_2Dim_()

    ' Arrange
    Dim a2, aex As New ArrEx
    a2 = array2d_1to3x0to12

    ' Act
    Set aex = aex.Initialize(a2).SelectColumn(2)

    ' Assert
    With UnitTest.NameOf("SelectColumn_2Dim_")
        Call .AssertEqual(1, aex.Rank)
        Call .AssertEqual(3, aex.Value(1))
        Call .AssertEqual(5, aex.Value(3))
    End With

End Sub

'[Fact]
Sub SelectColumn_01Dim_()

    ' Arrange
    Dim a0, a1, aex0 As New ArrEx, aex1 As New ArrEx
    a0 = 1
    a1 = array1d_1to10

    ' Act
    Set aex0 = aex0.Initialize(a0).SelectColumn(2)
    Set aex1 = aex1.Initialize(a1).SelectColumn(2)

    ' Assert
    With UnitTest.NameOf("SelectColumn_1Dim_")
        Call .AssertEqual(a0, aex0.Value)
        Call .AssertEqual(a1, aex1.Value)
    End With

End Sub

'[Fact]
Sub SelectRow_2Dim_()

    ' Arrange
    Dim a2, aex As New ArrEx
    a2 = array2d_1to3x0to12
    
    ' Act
    Set aex = aex.Initialize(a2).SelectRow(2)

    ' Assert
    With UnitTest.NameOf("SelectRow_2Dim_")
        Call .AssertEqual(1, aex.Rank)
        Call .AssertEqual(2, aex.Value(0))
        Call .AssertEqual(14, aex.Value(12))
    End With
    
End Sub

'[Fact]
Sub SelectRow_01Dim_()

    ' Arrange
    Dim a0, a1, aex0 As New ArrEx, aex1 As New ArrEx
    a0 = 1
    a1 = array1d_1to10

    ' Act
    Set aex0 = aex0.Initialize(a0).SelectRow(2)
    Set aex1 = aex1.Initialize(a1).SelectRow(2)

    ' Assert
    With UnitTest.NameOf("SelectRow_1Dim_")
        Call .AssertEqual(a0, aex0.Value)
        Call .AssertEqual(a1, aex1.Value)
    End With

End Sub

'[Fact]
Sub Where_2Dim_()

    ' Arrange
    Dim a2, aex As New ArrEx
    a2 = array2d_1to3x0to12
    Call aex.Initialize(a2)

    ' Act & Assert
    With UnitTest.NameOf("Where_2Dim_")
        Call .AssertEqual(1, aex.WhereBy(2, aexEqual, 3).Count)
        Call .AssertEqual(2, aex.WhereBy(2, aexDoesNotEqual, 3).Count)
        Call .AssertEqual(2, aex.WhereBy(2, aexGreaterThan, 3).Count)
        Call .AssertEqual(3, aex.WhereBy(2, aexGreaterThanOrEqualTo, 3).Count)
        Call .AssertEqual(1, aex.WhereBy(2, aexLessThan, 4).Count)
        Call .AssertEqual(2, aex.WhereBy(2, aexLessThanOrEqualTo, 4).Count)
        Call .AssertEqual(0, aex.WhereBy(2, aexEqual, 1000).Count)
    End With
    
End Sub

'[Fact]
Sub Where_1Dim_()

    ' Arrange
    Dim a, aex As New ArrEx
    a = array1d_1to10
    Call aex.Initialize(a)

    ' Act & Assert
    With UnitTest.NameOf("Where_1Dim_")
        Call .AssertEqual(1, aex.WhereBy(, aexEqual, 3).Count)
        Call .AssertEqual(9, aex.WhereBy(, aexDoesNotEqual, 3).Count)
        Call .AssertEqual(7, aex.WhereBy(, aexGreaterThan, 3).Count)
        Call .AssertEqual(8, aex.WhereBy(, aexGreaterThanOrEqualTo, 3).Count)
        Call .AssertEqual(3, aex.WhereBy(, aexLessThan, 4).Count)
        Call .AssertEqual(4, aex.WhereBy(, aexLessThanOrEqualTo, 4).Count)
        Call .AssertEqual(0, aex.WhereBy(, aexEqual, 1000).Count)
    End With
    
End Sub

'[Fact]
Sub FindIndex_2Dim_()

    ' Arrange
    Dim a2, aex As New ArrEx
    a2 = array2d_1to3x0to12
    Call aex.Initialize(a2)

    ' Act & Assert
    With UnitTest.NameOf("FindIndex_2Dim_")
        Call .AssertEqual(1, aex.FindIndex(2, aexEqual, 3))
        Call .AssertEqual(2, aex.FindIndex(2, aexDoesNotEqual, 3))
        Call .AssertEqual(2, aex.FindIndex(2, aexGreaterThan, 3))
        Call .AssertEqual(1, aex.FindIndex(2, aexGreaterThanOrEqualTo, 3))
        Call .AssertEqual(1, aex.FindIndex(2, aexLessThan, 4))
        Call .AssertEqual(1, aex.FindIndex(2, aexLessThanOrEqualTo, 4))
        Call .AssertEqual(-1, aex.FindIndex(2, aexEqual, 1000))
    End With
    
End Sub

'[Fact]
Sub FindIndex_1Dim_()

    ' Arrange
    Dim a, aex As New ArrEx
    a = array1d_1to10
    Call aex.Initialize(a)

    ' Act & Assert
    With UnitTest.NameOf("FindIndex_1Dim_")
        Call .AssertEqual(3, aex.FindIndex(, aexEqual, 3))
        Call .AssertEqual(2, aex.FindIndex(, aexDoesNotEqual, 1))
        Call .AssertEqual(4, aex.FindIndex(, aexGreaterThan, 3))
        Call .AssertEqual(3, aex.FindIndex(, aexGreaterThanOrEqualTo, 3))
        Call .AssertEqual(1, aex.FindIndex(, aexLessThan, 4))
        Call .AssertEqual(1, aex.FindIndex(, aexLessThanOrEqualTo, 4))
        Call .AssertEqual(-1, aex.FindIndex(, aexEqual, 1000))
    End With
    
End Sub


'[Fact]
Sub Skip_1Dim_()

    ' Arrange
    Dim a, aex As New ArrEx
    a = array1d_1to10
    Call aex.Initialize(a)

    ' Act & Assert
    With UnitTest.NameOf("Skip_1Dim_")
        Call .AssertEqual(7, aex.Skip(3).Count)
        Call .AssertEqual(a(4), aex.Skip(3).Value(1))
        Call .AssertEqual(0, aex.Skip(10).Count)
    End With
    
End Sub

'[Fact]
Sub Skip_2Dim_()

    ' Arrange
    Dim a, aex As New ArrEx
    a = array2d_1to3x0to12
    Call aex.Initialize(a)

    ' Act & Assert
    With UnitTest.NameOf("Skip_2Dim_")
        Call .AssertEqual(2, aex.Skip(1).Count)
        Call .AssertEqual(a(2, 0), aex.Skip(1).Value(1, 0))
        Call .AssertEqual(0, aex.Skip(10).Count)
                        
        Call .AssertEqual(10, aex.Skip(1, 3).Count(2))
        
    End With
    
End Sub

'[Fact]
Sub Take_2Dim_()
    
' Arrange
    Dim a, aex As New ArrEx
    a = array2d_1to3x0to12
    Call aex.Initialize(a)

    ' Act & Assert
    With UnitTest.NameOf("Take_2Dim_")
        Call .AssertEqual(2, aex.Take(2).Count)
        Call .AssertEqual(a(1, 0), aex.Take(2).Value(1, 0))

        Call .AssertEqual(0, aex.Take(0).Count)
        Call .AssertEqual(3, aex.Take(10).Count)
                        
        Call .AssertEqual(3, aex.Take(1, 3).Count(2))
    End With
        
End Sub

'[Fact]
Sub OrderBy_1Dim_Number()
    ' Arrange
    Dim a, aex As New ArrEx
    a = array1d_1to10_UnSort
    Call aex.Initialize(a)
    
    ' Act & Assert
    With UnitTest.NameOf("OrderBy")
        Call .AssertEqual("1,2,3,4,5,6,7,8,9,10", aex.OrderBy().ToString())
        Call .AssertEqual(1, aex.OrderBy().Value(1))
        Call .AssertEqual(10, aex.OrderBy().Value(10))
    End With
End Sub
    
'[Fact]
Sub OrderBy_1Dim_String()
    ' Arrange
    Dim a, aex As New ArrEx
    a = array1d_1to3_String
    Call aex.Initialize(a)
    
    ' Act & Assert
    With UnitTest.NameOf("OrderBy_String")
        Call .AssertEqual("aaa,aab,abc", aex.OrderBy().ToString())
    End With
End Sub

'[Fact]
Sub OrderByDescending_1Dim_Number()
    ' Arrange
    Dim a, aex As New ArrEx
    a = array1d_1to10_UnSort
    Call aex.Initialize(a)
    
    ' Act & Assert
    With UnitTest.NameOf("OrderByDescending")
        Call .AssertEqual("10,9,8,7,6,5,4,3,2,1", aex.OrderByDescending().ToString())
        Call .AssertEqual(10, aex.OrderByDescending().Value(1))
        Call .AssertEqual(1, aex.OrderByDescending().Value(10))
    End With
End Sub
    
'[Fact]
Sub OrderByDescending_1Dim_String()
    ' Arrange
    Dim a, aex As New ArrEx
    a = array1d_1to3_String
    Call aex.Initialize(a)
    
    ' Act & Assert
    With UnitTest.NameOf("OrderByDescending_String")
        Call .AssertEqual("abc,aab,aaa", aex.OrderByDescending().ToString())
    End With
End Sub
    
'[Fact]
Sub OrderBy_2Dim()
    ' Arrange
    Dim a, aex As New ArrEx
    a = array2d_1to3x1to2_UnSort
    Call aex.Initialize(a)
    
    ' Act & Assert
    With UnitTest.NameOf("OrderBy_2D")
        Call .AssertEqual("1,1;2,2;3,3", aex.OrderBy(1).ToString())
        Call .AssertEqual(1, aex.OrderBy(1).Value(1, 1))
        Call .AssertEqual(3, aex.OrderBy(1).Value(3, 2))
    End With
    
End Sub
    
'[Fact]
Sub OrderByDescending_2Dim()
    ' Arrange
    Dim a, aex As New ArrEx
    a = array2d_1to3x1to2_UnSort
    Call aex.Initialize(a)
    
    ' Act & Assert
    With UnitTest.NameOf("OrderByDescending_2D")
        Call .AssertEqual("3,3;2,2;1,1", aex.OrderByDescending(2).ToString())
        Call .AssertEqual(3, aex.OrderByDescending(2).Value(1, 1))
        Call .AssertEqual(1, aex.OrderByDescending(2).Value(3, 2))
    End With
    
End Sub
    
'[Fact]
Sub Distinct_0Dim_()
    ' Arrange
    Dim a, aex As New ArrEx
    a = 1
    Call aex.Initialize(a)
    
    ' Act & Assert
    With UnitTest.NameOf("Distinct - 0 dim")
        Call .AssertEqual("1", aex.Distinct().ToString())
    End With
End Sub
    
'[Fact]
Sub Distinct_1Dim_()
    ' Arrange
    Dim a, aex As New ArrEx
    a = array1d_1to10_HasDuplicationValue
    Call aex.Initialize(a)
    
    ' Act & Assert
    With UnitTest.NameOf("Distinct - 1 dim")
        Call .AssertEqual("1,2,5,6,7,8,9", aex.Distinct().ToString())
        Call .AssertEqual("1,2,5,6,7,8,9", aex.Cast(vbString).Distinct().ToString())
    End With
End Sub
    
'[Fact]
Sub Distinct_2Dim_()
    ' Arrange
    Dim a, aex As New ArrEx
    a = array2d_1to3x1to2_Duplication
    Call aex.Initialize(a)
    
    ' Act & Assert
    With UnitTest.NameOf("Distinct - 2 dim")
        Call .AssertEqual("1,2;2,2", aex.Distinct().ToString())
    End With
End Sub
    
'[Fact]
Sub DistinctBy_2Dim_()
    ' Arrange
    Dim a, aex As New ArrEx
    a = array2d_1to3x1to2_Duplication
    Call aex.Initialize(a)
    
    ' Act & Assert
    With UnitTest.NameOf("DistinctBy - 2 dim")
        Call .AssertEqual("1,2", aex.DistinctBy(2).ToString())
    End With
End Sub
    
'[Fact]
Sub Transpose__()
    ' Arrange
    Dim aex As New ArrEx
        
    ' Act & Assert
    With UnitTest.NameOf("Transpose")
        Call .AssertEqual(1, aex.Create(1).Transpose().Value)
        Call .AssertEqual("1;2;3;4;5;6;7;8;9;10", aex.Create(array1d_1to10).Transpose().ToString)
        Call .AssertEqual(1, aex.Create(array1d_1to10).Transpose().LowerBound(1))
        Call .AssertEqual(10, aex.Create(array1d_1to10).Transpose().UpperBound(1))
        Call .AssertEqual(0, aex.Create(array1d_1to10).Transpose().UpperBound(2))
        Call .AssertEqual("1,4;2,5;3,6", aex.Create(array2d_2x3).Transpose().ToString())
        Call .AssertEqual(array2d_2x3, aex.Create(array2d_2x3).Transpose().Transpose().Value)
    End With
    
End Sub
    
    
'[Fact]
Sub XLookUp_1Dim_()
    ' Arrange
    Dim a, aex As New ArrEx
    a = array1d_1to10
    Call aex.Initialize(a)
    
    ' Act & Assert
    With UnitTest.NameOf("XLookUp")
        Call .AssertEqual("3,4,5,0,0,0,0,0,0,0", aex.XLookUp(array2d_1to3x0to12, 2, 0, 0).ToString())
        Call .AssertEqual("3,4,5,,,,,,,", aex.XLookUp(array2d_1to3x0to12, 2, 0).ToString())
    End With
End Sub
'[Fact]
Sub XLookUp_2Dim_()
    ' Arrange
    Dim a, aex As New ArrEx
    a = array2d_2x3
    Call aex.Initialize(a)
    
    ' Act & Assert
    With UnitTest.NameOf("XLookUp 2 dim")
        Call .AssertEqual("3,4,5;0,0,0", aex.XLookUp(array2d_1to3x0to12, 2, 0, 0).ToString())
    End With
End Sub

'[Fact]
Sub Cast__()
    ' Arrange
    Dim aex As New ArrEx
    
    ' Act & Assert
    With UnitTest.NameOf("Cast")
        Call .AssertEqual(vbBoolean, VarType(aex.Create(1).Cast(vbBoolean).Value))
        Call .AssertEqual(vbCurrency, VarType(aex.Create(1).Cast(vbCurrency).Value))
        Call .AssertEqual(vbString, VarType(aex.Create(1).Cast(vbString).Value))
        
        Call .AssertEqual(vbBoolean, VarType(aex.Create(array1d_1to10).Cast(vbBoolean).Value(10)))
        Call .AssertEqual(vbBoolean, VarType(aex.Create(array2d_2x3).Cast(vbBoolean).Value(2, 3)))
    End With
    
    
End Sub
    
'[Fact]
Sub ToCollection_1Dim_()
    ' Arrange
    Dim aex As New ArrEx, col As Collection
    Call aex.Initialize(array1d_1to10)
    
    ' Act
    Set col = aex.ToCollection()
    
    ' Assert
    With UnitTest.NameOf("ToCollection 1 dim")
        Call .AssertEqual(10, col.Count)
    End With
End Sub

'[Fact]
Sub ToCollection_2Dim_()
    ' Arrange
    Dim aex As New ArrEx, col As Collection
    Call aex.Initialize(array2d_1to3x0to12)
    
    ' Act & Assert
    With UnitTest.NameOf("ToCollection 2 dim horizontal")
        Set col = aex.ToCollection(aexHorizontal)
        Call .AssertEqual(3, col.Count)
        Call .AssertEqual(0, LBound(col(1)))
        Call .AssertEqual(12, UBound(col(1)))
    End With

    With UnitTest.NameOf("ToCollection 2 dim vertical")
        Set col = aex.ToCollection(aexVertical)
        Call .AssertEqual(13, col.Count)
        Call .AssertEqual(1, LBound(col(1)))
        Call .AssertEqual(3, UBound(col(1)))
    End With

End Sub


'[Fact]
Sub InnerJoin__()
    
    ' Arrange
    Dim aex As New ArrEx
    
    ' Act
    Set aex = ArrEx(array_Persons).InnerJoin(array_Scores, 0, 0)

    ' Assert
    With UnitTest.NameOf("Inner Join")
        Call .AssertEqual(3, aex.Count)
        Call .AssertEqual("ID,Name,ID,Score;2,Bob,2,85;3,Charlie,3,90", aex.ToString())
    End With

End Sub

'[Fact]
Sub LeftJoin__()
    
    ' Arrange
    Dim aex As New ArrEx
    
    ' Act
    Set aex = ArrEx(array_Persons).LeftJoin(array_Scores, 0, 0)

    ' Assert
    With UnitTest.NameOf("Left Join")
        Call .AssertEqual(4, aex.Count)
        Call .AssertEqual("ID,Name,ID,Score;1,Alice,,;2,Bob,2,85;3,Charlie,3,90", aex.ToString())
    End With

End Sub

'[Fact]
Sub FullOuterJoin__()
    
    ' Arrange
    Dim aex As New ArrEx
    
    ' Act
    Set aex = ArrEx(array_Persons).FullOuterJoin(array_Scores, 0, 0)

    ' Assert
    With UnitTest.NameOf("Full Outer Join")
        Call .AssertEqual(5, aex.Count)
        Call .AssertEqual("ID,Name,ID,Score;1,Alice,,;2,Bob,2,85;3,Charlie,3,90;,,4,78", aex.ToString())
    End With

End Sub


'[Fact]
Sub CrossJoin__()
    
    ' Arrange
    Dim aex As New ArrEx
    
    ' Act
    Set aex = ArrEx(array_Persons).CrossJoin(array_Scores)

    ' Assert
    With UnitTest.NameOf("Cross Join")
        Call .AssertEqual((UBound(array_Persons) + 1) * (UBound(array_Scores) + 1), aex.Count)
    End With

End Sub

    
'Fact]
Sub Test__()
    ' Arrange
    ' Act
    ' Assert
End Sub


















Private Function GetArray1d() As Variant
    Dim a
    ReDim a(1 To 5)
    Dim i
    For i = 1 To 5
        a(i) = i
    Next i
    GetArray1d = a
End Function

Private Function array1d_1to10() As Variant
    Dim a
    ReDim a(1 To 10)
    Dim i
    For i = 1 To 10
        a(i) = i
    Next i
    array1d_1to10 = a
End Function

Private Function array1d_1to10_UnSort() As Variant
    Dim a
    ReDim a(1 To 10)
    a(1) = 4
    a(2) = 3
    a(3) = 2
    a(4) = 1
    a(5) = 10
    a(6) = 9
    a(7) = 8
    a(8) = 7
    a(9) = 6
    a(10) = 5
    array1d_1to10_UnSort = a
End Function

Private Function array1d_1to3_String() As Variant
    Dim a
    ReDim a(1 To 3)
    a(1) = "aaa"
    a(2) = "abc"
    a(3) = "aab"
    array1d_1to3_String = a
End Function

Private Function array1d_1to10_HasDuplicationValue() As Variant
    Dim a
    ReDim a(1 To 10)
    a(1) = 1
    a(2) = 2
    a(3) = 2
    a(4) = 2
    a(5) = 5
    a(6) = 6
    a(7) = 7
    a(8) = 8
    a(9) = 9
    a(10) = 6
    array1d_1to10_HasDuplicationValue = a
End Function

Private Function GetArray2d() As Variant
    Dim a
    ReDim a(1 To 5, 1 To 10)
    Dim i, j
    For i = 1 To 5
        For j = 1 To 10
            a(i, j) = i * j
        Next j
    Next i
    GetArray2d = a
End Function

Private Function array2d_1to3x0to12() As Variant
    Dim a
    ReDim a(1 To 3, 0 To 12)
    Dim i, j
    For i = 1 To 3
        For j = 0 To 12
            a(i, j) = i + j
        Next j
    Next i
    array2d_1to3x0to12 = a
End Function

Private Function array2d_2x3() As Variant
    Dim a, cnt As Long
    ReDim a(1 To 2, 1 To 3)
    Dim i, j
    For i = 1 To 2
        For j = 1 To 3
            cnt = cnt + 1
            a(i, j) = cnt
        Next j
    Next i
    array2d_2x3 = a
End Function

Private Function array2d_1to3x1to2_UnSort() As Variant
    Dim a
    ReDim a(1 To 3, 1 To 2)
    a(1, 1) = 3
    a(1, 2) = 3
    a(2, 1) = 1
    a(2, 2) = 1
    a(3, 1) = 2
    a(3, 2) = 2
    array2d_1to3x1to2_UnSort = a
End Function


Private Function array2d_1to3x1to2_Duplication() As Variant
    Dim a
    ReDim a(1 To 3, 1 To 2)
    a(1, 1) = 1
    a(1, 2) = 2
    a(2, 1) = 2
    a(2, 2) = 2
    a(3, 1) = 1
    a(3, 2) = 2
    array2d_1to3x1to2_Duplication = a
End Function

Private Function array_Persons()
    Dim arr1(0 To 3, 0 To 1)
    
    arr1(0, 0) = "ID"
    arr1(0, 1) = "Name"
    arr1(1, 0) = 1
    arr1(1, 1) = "Alice"
    arr1(2, 0) = 2
    arr1(2, 1) = "Bob"
    arr1(3, 0) = 3
    arr1(3, 1) = "Charlie"
    
    array_Persons = arr1
End Function

Private Function array_Scores()
    Dim arr2(0 To 3, 0 To 1)
    
    arr2(0, 0) = "ID"
    arr2(0, 1) = "Score"
    arr2(1, 0) = 2
    arr2(1, 1) = 85
    arr2(2, 0) = 3
    arr2(2, 1) = 90
    arr2(3, 0) = 4
    arr2(3, 1) = 78
    
    array_Scores = arr2
End Function

