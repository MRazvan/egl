VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Vec4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public x As Double
Public y As Double
Public z As Double
Public w As Double

Public Sub Clear()
    SetValue 0
End Sub

Public Sub SetValue(val As Double)
    x = val
    y = val
    z = val
    w = val
End Sub

Public Sub Copy(ByRef from As Vec4)
    x = from.x
    y = from.y
    z = from.z
    w = from.w
End Sub

Public Sub Normalize()
    Dim mag As Double
    mag = Math.Sqr(x * x + y * y + z * z + w * w)
    If mag <> 0 Then
        x = x / mag
        y = y / mag
        z = z / mag
        w = w / mag
    End If
End Sub

Public Sub Clamp(min As Double, max As Double)
    If x < min Then
        x = min
    ElseIf x > max Then
        x = max
    End If
    If y < min Then
        y = min
    ElseIf y > max Then
        y = max
    End If
    If z < min Then
        z = min
    ElseIf z > max Then
        z = max
    End If
    If w < min Then
        w = min
    ElseIf w > max Then
        w = max
    End If
End Sub
