Attribute VB_Name = "Test"
Option Explicit

'@Test
Public Function TestJoin2() As Boolean
    Dim X As New Vector, Y As New Vector, Z As New Vector
    X.Contents = Array(1, 2, 3, 4)
    Y.Contents = Array(5, 6, 7, 8)
    Z.Contents = Array(1, 2, 3, 4, 5, 6, 7, 8)
    TestJoin2 = Identical(Join(X, Y), Z)
End Function


'@Test
Public Function TestJoin3() As Boolean
    Dim X As New Vector, Y As New Vector, Z As New Vector, E As New Vector
    X.Contents = Array(1, 2, 3, 4)
    Y.Contents = Array(5, 6, 7, 8)
    Z.Contents = Array(9, 10)
    E.Contents = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10)
    TestJoin3 = Identical(Join(X, Y, Z), E)
End Function


'@Test
Public Function TestShow() As Boolean
    Dim M As New Matrix, Contents, i As Integer, j As Integer
    ReDim Contents(1 To 6, 1 To 3)
    For i = 1 To 6
        For j = 1 To 3
            Contents(i, j) = i * j
        Next
    Next
    M.Contents = Contents
    Debug.Print M.Show
End Function


'@Test
Public Function TestShow2() As Boolean
    Dim M As New Matrix, Contents, i As Integer, j As Integer
    ReDim Contents(1 To 6, 1 To 3)
    For i = 1 To 6
        For j = 1 To 3
            Contents(i, j) = i * j
        Next
    Next
    M.Contents = Contents
    Debug.Print Show(M)
    Debug.Print Transpose(M).Show
End Function
