Attribute VB_Name = "Misc"
Option Explicit

Public Function Dimensions(Arr As Variant) As Integer
    Dim dummy As Integer, d As Integer
    On Error GoTo Catch
    Do While True
        d = d + 1
        dummy = UBound(Arr, d)
    Loop
    Exit Function
Catch:
    Dimensions = d - 1
End Function

Public Function Pad(s As String, Width As Byte) As String
    Dim padding As String, counter As Byte
    For counter = Len(s) To Width
        padding = padding & " "
    Next
    Pad = s & Left(padding, Len(padding) - 1)
End Function
