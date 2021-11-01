Attribute VB_Name = "Functions"
Option Explicit

Public Function Join(ParamArray Vectors() As Variant) As Vector
    Dim V, Values, Length As Long, i As Long, j As Long
    For Each V In Vectors
        Length = Length + Size(V)
    Next
    
    ReDim Values(1 To Length)
    For Each V In Vectors
        For i = 1 To Size(V)
            j = j + 1
            Values(j) = V.Ix(i)
        Next
    Next
    
    Set Join = New Vector
    Join.Contents = Values
    
End Function

Public Function Size(ByVal V As Vector) As Long
    Size = V.Length
End Function

Public Function Identical(X As Vector, Y As Vector) As Boolean
    If Size(X) <> Size(Y) Then Exit Function
    Dim i As Long
    For i = LBound(X.Contents) To UBound(X.Contents)
        If X.Contents(i) <> Y.Contents(i) Then Exit Function
    Next
    Identical = True
End Function

Public Function Show(Obj As Variant) As String
    If IsObject(Obj) Then
        Show = Obj.Show
    End If
End Function

Public Function Transpose(M As Matrix) As Matrix
    Dim i As Long, j As Long
    Set Transpose = New Matrix
    Transpose.Resize M.Ncol, M.Nrow
    For i = 1 To M.Ncol
        For j = 1 To M.Nrow
            Transpose.Mutate i, j, M.Ix(j, i)
        Next
    Next
End Function
