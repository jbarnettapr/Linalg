VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Vector"
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
    If Dimensions(Values) <> 1 Then Err.Raise 1
    pContents = Values
End Property

Public Function Ix(i As Long)
    Idx = pContents(LBound(pContents) + i - 1)
End Function

Public Function Show() As String
    Dim i As Long
    Show = "<"
    For i = 1 To Me.Length
        Show = Show & Me.Ix(i) & ", "
    Next
    Show = Left(Show, Len(Show) - 2) & ">"
End Function

Public Property Get Length() As Long
    Length = 1 + UBound(pContents) - LBound(pContents)
End Property

Public Sub Resize(Length As Long)
    ReDim Preserve pContents(1 To Length)
End Sub

