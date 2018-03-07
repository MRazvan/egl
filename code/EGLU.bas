Attribute VB_Name = "EGLU"
' http://www.songho.ca/opengl/gl_matrix.html

Private Sub SetPerspectiveFrustrum(ByRef mat As Matrix, left As Double, right As Double, bottom As Double, top As Double, near As Double, far As Double)
    mat.MatVal(0, 0) = (2 * near) / (right - left)
    mat.MatVal(1, 1) = (2 * near) / (top - bottom)
    mat.MatVal(2, 0) = (right + left) / (right - left)
    mat.MatVal(2, 1) = (top + bottom) / (top - bottom)
    mat.MatVal(2, 2) = -(far + near) / (far - near)
    mat.MatVal(2, 3) = -1
    mat.MatVal(3, 2) = -(2 * far * near) / (far - near)
    mat.MatVal(3, 3) = 0
End Sub

Public Sub gluPerspective(fovy As Double, aspect As Double, near As Double, far As Double)

    Dim height As Double
    Dim width As Double
    
    height = Math.Tan(WorksheetFunction.Radians(fovy) / 2) * near
    width = height * aspect
    
    Dim res As New Matrix
    SetPerspectiveFrustrum res, -width, width, -height, height, near, far
    
    EGL.gLoadMatrix res
    
End Sub

Public Sub gluOrtho(left As Double, right As Double, bottom As Double, top As Double, near As Double, far As Double)
    Dim res As New Matrix
    res.MatVal(0, 0) = 2 / (right - left)
    res.MatVal(1, 1) = 2 / (top - bottom)
    res.MatVal(2, 2) = -2 / (far - near)
    res.MatVal(3, 0) = -(right + left) / (right - left)
    res.MatVal(3, 1) = -(top + bottom) / (top - bottom)
    res.MatVal(3, 2) = -(far + near) / (far - near)
    
    EGL.gLoadMatrix res
End Sub
