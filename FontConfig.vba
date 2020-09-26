Option Explicit

Public name As String
Public size As Double
Public expanded As Double
Public probability As Double

Private Sub Class_Initialize()
    name = vbNullString
    size = 0
    expanded = 0
    probability = 0
End Sub

Public Sub InitializeWithValues(ByVal name_ As String, ByVal size_ As Double, ByVal expanded_ As Double, ByVal probability_ As Double)
    name = name_
    size = size_
    expanded = expanded_
    probability = probability_
End Sub

Public Sub InitializeDefaultValues()
    name = vbNullString
    size = 0
    expanded = 0
    probability = 0
End Sub
