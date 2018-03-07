VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Initialize()
    TestRenderer.Initialize 6, 128, 64, "Display"
End Sub

Sub TestTriangle()
    Initialize
    While True
        TestRenderer.ModelRotation.z = TestRenderer.ModelRotation.z + 10
        TestRenderer.RunTriangle
    Wend
End Sub

Sub TestCubeWithQuads()
    Initialize
    While True
        TestRenderer.ModelRotation.x = TestRenderer.ModelRotation.x + 10
        TestRenderer.ModelRotation.y = TestRenderer.ModelRotation.y + 10
        TestRenderer.RunQuads
    Wend
End Sub

Sub TestColoredTriangle()
    Initialize
    While True
        TestRenderer.ModelRotation.z = TestRenderer.ModelRotation.z + 10
        TestRenderer.RunSingleTriangle
    Wend
End Sub

Sub TestCubeWithTriangles()
    Initialize
    While True
        TestRenderer.ModelRotation.x = TestRenderer.ModelRotation.x + 10
        TestRenderer.ModelRotation.y = TestRenderer.ModelRotation.y + 10
        TestRenderer.RunTriangles True
    Wend
End Sub



