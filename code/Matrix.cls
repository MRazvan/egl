VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Matrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Data As Variant

Public Property Let MatVal(row As Integer, col As Integer, val As Double)
    Data(row, col) = val
End Property


Private Sub Class_Initialize()
    Dim arr(3, 3) As Double
    Data = arr
End Sub

Public Function Transpose() As Matrix
    Dim result As New Matrix
    result.Data = WorksheetFunction.Transpose(Data)
    Set Transpose = result
End Function


 
