VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BoundingBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public min As Vec3
Public max As Vec3

Private Sub Class_Initialize()
    Set min = New Vec3
    min.SetValue 1000000
    Set max = New Vec3
    max.Clear
End Sub

