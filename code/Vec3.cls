VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Vec3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public x As Double
Public y As Double
Public z As Double

Public Sub Clear()
    SetValue 0
End Sub

Public Sub SetValue(val As Double)
    x = val
    y = val
    z = val
End Sub

Public Sub Copy(ByRef from As Vec3)
    x = from.x
    y = from.y
    z = from.z
End Sub

