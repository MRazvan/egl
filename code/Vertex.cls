VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Vertex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Position As Vec4
Public Normal As Vec4
Public VertColor As Vec4

Private Sub Class_Initialize()
    Set Position = New Vec4
    Set Normal = New Vec4
    Set VertColor = New Vec4
End Sub
