VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransformValues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Translate As Vec3
Public Scaling As Vec3
Public Rotation As Vec3

Public SinRotation As Vec3
Public CosRotation As Vec3

Private Sub Class_Initialize()
    Set Translate = New Vec3
    Set Scaling = New Vec3
    Set Rotation = New Vec3

    Set SinRotation = New Vec3
    Set CosRotation = New Vec3
    
    Scaling.x = 1
    Scaling.y = 1
    Scaling.z = 1
End Sub

Public Sub CalculateRotationValues()
    SinRotation.x = Math.Sin(WorksheetFunction.Radians(Rotation.x))
    SinRotation.y = Math.Sin(WorksheetFunction.Radians(Rotation.y))
    SinRotation.z = Math.Sin(WorksheetFunction.Radians(Rotation.z))
    
    CosRotation.x = Math.Cos(WorksheetFunction.Radians(Rotation.x))
    CosRotation.y = Math.Cos(WorksheetFunction.Radians(Rotation.y))
    CosRotation.z = Math.Cos(WorksheetFunction.Radians(Rotation.z))
End Sub


