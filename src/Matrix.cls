VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Matrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'@Expose

Private pContents As Variant

'@Default
Public Property Get Contents()
    Contents = pContents
End Property

Public Property Let Contents(Values As Variant)
    If Dimensions(Values) <> 2 Then Err.Raise 1
    pContents = Values
End Property

Public Function Show() As String
    Dim i As Long, j As Long, colWidths() As Byte, elem, elemWidth As Byte
    
    ' One pass to get column widths for formatting
    ReDim colWidths(1 To Me.Ncol)
    For j = 1 To Me.Ncol
        colWidths(j) = 0
        For i = 1 To Me.Nrow
            elem = Me.Ix(i, j)
            elemWidth = Len(CStr(elem))
            If elemWidth > colWidths(j) Then colWidths(j) = elemWidth
        Next
    Next
    
    ' Now construct the string
    For i = 1 To Me.Nrow
        Show = Show & "| "
        For j = 1 To Me.Ncol
            Show = Show & Pad(Me.Ix(i, j), colWidths(j)) & "  "
        Next
        Show = Show & "|" & vbNewLine
    Next
    
End Function

Public Function Ix(i As Long, j As Long)
    Ix = pContents(LBound(pContents, 1) + i - 1, LBound(pContents, 2) + j - 1)
End Function

Public Sub Mutate(ByVal i As Long, ByVal j As Long, Value)
    On Error GoTo Catch
    pContents(LBound(pContents, 1) + i - 1, LBound(pContents, 2) + j - 1) = Value
    Exit Sub
Catch:
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Property Get Nrow() As Long
    Nrow = 1 + UBound(pContents, 1) - LBound(pContents, 1)
End Property

Public Property Get Ncol() As Long
    Ncol = 1 + UBound(pContents, 2) - LBound(pContents, 2)
End Property

Public Property Get Size() As Long
    Size = (1 + UBound(pContents, 1) - LBound(pContents, 1)) * (1 + UBound(pContents, 2) - LBound(pContents, 2))
End Property

Public Sub Resize(Length As Long, Width As Long)
    If Length = 0 Or Width = 0 Then pContents = Empty: Exit Sub
    If IsEmpty(pContents) Then ReDim pContents(1 To 1, 1 To 1)
    ReDim Preserve pContents(LBound(pContents, 1) To UBound(pContents, 1), LBound(pContents, 2) To LBound(pContents, 2) + Width - 1)
    TransposeMe
    ReDim Preserve pContents(LBound(pContents, 1) To UBound(pContents, 1), LBound(pContents, 2) To LBound(pContents, 2) + Length - 1)
    TransposeMe
End Sub

Private Sub TransposeMe()
    Dim i As Long, j As Long, X()
    ReDim Preserve X(1 To Ncol, 1 To Nrow)
    For i = 1 To Nrow
        For j = 1 To Ncol
            X(j, i) = Ix(i, j)
        Next
    Next
    Me.Contents = X
End Sub
