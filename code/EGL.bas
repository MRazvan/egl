Attribute VB_Name = "EGL"
Enum gBeginEnum
    None = 0
    Points = 1
    Lines = 2
    LineStrips = 3
    Triangle = 4
    TriangleStrips = 5
    Quads = 6
    QuadStrips = 7
End Enum

Enum gPolygonModeEnum
    Line = 1
    Fill = 2
End Enum

Enum gClearEnum
    ColorBit = 1
    DepthBit = 2
End Enum

Enum gMatrixModeEnum
    model = 0
    projection = 1
End Enum

' Matrix state
Private IdentityMatrix As Matrix
Private MatrixMode As gMatrixModeEnum
Private Matrices() As Matrix
Private ModelTransform As TransformValues

'
Private Surface As Display
Private HalfSurfaceWidth As Double
Private HalfSurfaceHeight As Double

' Draw mode
Private CurrentState As gBeginEnum
Private CurrentPolygonMode As gPolygonModeEnum

Private ClearColor As Vec4
Private CurrentColor As Vec4

' Various buffers
Private VertexBuffer() As Vertex
Private PrimaryBuffer() As Long
Private DepthBuffer() As Double
Private BufferSize As Long
Private CurrentNormal As Vec4

Private MVPMatrix As Matrix

' The number of vertex elements we have
Private CurrentIdx As Integer

' Shader methods
Private VertexShaderMethodName As String
Private FragmentShaderMethodName As String

' Keep this as cache
Private Frag As Fragment

Sub gSetVertexShader(vertexSMethodName As String)
    VertexShaderMethodName = vertexSMethodName
End Sub

Sub gSetFragmentShader(fragmentSMethodName As String)
    FragmentShaderMethodName = fragmentSMethodName
End Sub

Sub gPolygonMode(polyMode As gPolygonModeEnum)
    CurrentPolygonMode = polyMode
End Sub

Sub gInitialize(ByRef targetSurface As Display)
    Set Surface = targetSurface
    HalfSurfaceWidth = Surface.width / 2#
    HalfSurfaceHeight = Surface.height / 2#
    
    ReDim VertexBuffer(65536)
    BufferSize = CLng(Surface.width) * CLng(Surface.height)
    ReDim PrimaryBuffer(BufferSize)
    ReDim DepthBuffer(BufferSize)
    
    ' Redimension matrices
    ReDim Matrices(2)
    Set Matrices(gMatrixModeEnum.model) = New Matrix
    Set Matrices(gMatrixModeEnum.projection) = New Matrix
    ' Initialize the matrices
    Utils.IdentMatrix Matrices(gMatrixModeEnum.model)
    Utils.IdentMatrix Matrices(gMatrixModeEnum.projection)
    
    Set IdentityMatrix = New Matrix
    Utils.IdentMatrix IdentityMatrix
    
    Set ModelTransform = New TransformValues
    
    Set Frag = New Fragment
    
    Set ClearColor = New Vec4
    Set CurrentColor = New Vec4
    Set CurrentNormal = New Vec4
    
    gClearColor 0, 0, 0
    gClear ColorBit
    gPolygonMode gPolygonModeEnum.Fill
    PresentBuffer
End Sub

Sub gClearColor(r As Double, g As Double, b As Double)
    ClearColor.x = r
    ClearColor.y = g
    ClearColor.z = b
    ClearColor.Clamp 0, 1
End Sub

Sub gBegin(beginType As gBeginEnum)
    CurrentState = beginType
    CurrentIdx = 0
End Sub

Sub gClear(what As gClearEnum)
    Dim idx As Long
    If (what And gClearEnum.ColorBit) = ColorBit Then
        For idx = 1 To BufferSize
            PrimaryBuffer(idx) = RGB(ClearColor.x * 255, ClearColor.y * 255, ClearColor.z * 255)
        Next idx
    End If
    If (what And gClearEnum.DepthBit) = DepthBit Then
        For idx = 1 To BufferSize
            DepthBuffer(idx) = 10000000
        Next idx
    End If
End Sub

Sub gColor3b(r As Byte, g As Byte, b As Byte)
    CurrentColor.x = CDbl(r) / 255
    CurrentColor.y = CDbl(g) / 255
    CurrentColor.z = CDbl(b) / 255
    CurrentColor.Clamp 0, 1
End Sub

Sub gColor3d(r As Double, g As Double, b As Double)
    CurrentColor.x = r
    CurrentColor.y = g
    CurrentColor.z = b
    CurrentColor.Clamp 0, 1
End Sub

Sub gVertex2i(x As Integer, y As Integer)
    gVertex3i x, y, 1
End Sub

Sub gVertex3i(x As Integer, y As Integer, z As Integer)
    Dim v As New Vertex
    v.Position.x = x
    v.Position.y = y
    v.Position.z = z
    v.Position.w = 1#
    v.VertColor.Copy CurrentColor
    v.Normal.Copy CurrentNormal
    Set VertexBuffer(CurrentIdx) = v
    CurrentIdx = CurrentIdx + 1
End Sub

Sub gVertex2d(x As Double, y As Double)
    gVertex3d x, y, 1
End Sub

Sub gVertex3d(x As Double, y As Double, z As Double)
    Dim v As New Vertex
    v.Position.x = x
    v.Position.y = y
    v.Position.z = z
    v.Position.w = 1#
    v.VertColor.Copy CurrentColor
    v.Normal.Copy CurrentNormal
    Set VertexBuffer(CurrentIdx) = v
    CurrentIdx = CurrentIdx + 1
End Sub

Sub gMatrixMode(mode As gMatrixModeEnum)
    MatrixMode = mode
End Sub

Sub gLoadIdentity()
    Utils.IdentMatrix Matrices(MatrixMode)
End Sub

Sub gLoadMatrix(ByRef mat As Matrix)
    Set Matrices(MatrixMode) = mat
End Sub

Sub gTranslate(x As Double, y As Double, z As Double)
    With ModelTransform
        .Translate.x = x
        .Translate.y = y
        .Translate.z = z
    End With
End Sub

Sub gScale(x As Double, y As Double, z As Double)
    With ModelTransform
        .Scaling.x = x
        .Scaling.y = y
        .Scaling.z = z
    End With
End Sub

Sub gRotate(x As Double, y As Double, z As Double)
    With ModelTransform
        .Rotation.x = x
        .Rotation.y = y
        .Rotation.z = z
    End With
End Sub
Private Sub gDrawPointsInternal()
    Dim crtIdx As Integer
    For crtIdx = 0 To CurrentIdx - 1
        ProcessVertex VertexBuffer(crtIdx)
        SetPixel VertexBuffer(crtIdx)
    Next crtIdx
End Sub

Private Sub gDrawTriangleInternal()
    Dim crtIdx As Integer
    CurrentIdx = (CurrentIdx - (CurrentIdx Mod 3))
    For crtIdx = 0 To CurrentIdx - 1 Step 3
        ProcessVertex VertexBuffer(crtIdx)
        ProcessVertex VertexBuffer(crtIdx + 1)
        ProcessVertex VertexBuffer(crtIdx + 2)
        If CurrentPolygonMode = gPolygonModeEnum.Fill Then
            RasterTriangle VertexBuffer(crtIdx), VertexBuffer(crtIdx + 1), VertexBuffer(crtIdx + 2)
        Else
            DrawTriangle VertexBuffer(crtIdx), VertexBuffer(crtIdx + 1), VertexBuffer(crtIdx + 2)
        End If
    Next crtIdx
End Sub

Private Sub gDrawTriangleStripsInternal()
    Dim crtIdx As Integer
    For crtIdx = 0 To CurrentIdx - 3
        ProcessVertex VertexBuffer(crtIdx)
        ProcessVertex VertexBuffer(crtIdx + 1)
        ProcessVertex VertexBuffer(crtIdx + 2)
        If CurrentPolygonMode = gPolygonModeEnum.Fill Then
            RasterTriangle VertexBuffer(crtIdx), VertexBuffer(crtIdx + 1), VertexBuffer(crtIdx + 2)
        Else
            DrawTriangle VertexBuffer(crtIdx), VertexBuffer(crtIdx + 1), VertexBuffer(crtIdx + 2)
        End If
    Next crtIdx
End Sub

Private Sub gDrawQuadsInternal()
    Dim crtIdx As Integer
    For crtIdx = 0 To CurrentIdx - 1 Step 4
        ProcessVertex VertexBuffer(crtIdx)
        ProcessVertex VertexBuffer(crtIdx + 1)
        ProcessVertex VertexBuffer(crtIdx + 2)
        ProcessVertex VertexBuffer(crtIdx + 3)
        ' Call the vertex shader for all vertices
        DrawLine VertexBuffer(crtIdx), VertexBuffer(crtIdx + 1)
        DrawLine VertexBuffer(crtIdx + 1), VertexBuffer(crtIdx + 2)
        DrawLine VertexBuffer(crtIdx + 2), VertexBuffer(crtIdx + 3)
        DrawLine VertexBuffer(crtIdx + 3), VertexBuffer(crtIdx)
    Next crtIdx
End Sub


Sub gEnd()
    CalculateModelMatrix
    Set MVPMatrix = Utils.MatMult3(Matrices(gMatrixModeEnum.model), IdentityMatrix, Matrices(gMatrixModeEnum.projection))
    ' Now that we have the mvp matrix we need to multiply each vertex with that
    Select Case CurrentState
        Case gBeginEnum.Points
            gDrawPointsInternal
        Case gBeginEnum.Triangle
            gDrawTriangleInternal
        Case gBeginEnum.TriangleStrips
            gDrawTriangleStripsInternal
        Case gBeginEnum.Quads
            gDrawQuadsInternal
    End Select
End Sub

Public Sub gFlush()
    PresentBuffer
End Sub

Private Sub PresentBuffer()
    Surface.DisplayBuffer PrimaryBuffer
    Surface.SwapBuffers
    DoEvents
End Sub

Private Sub DrawLine(ByRef v1 As Vertex, ByRef v2 As Vertex)
    Dim x1I, y1I, x2I, y2I As Long
    x1I = v1.Position.x
    y1I = v1.Position.y
    x2I = v2.Position.x
    y2I = v2.Position.y

    Dim steep As Boolean
    steep = False
    
    Dim dx, dy, derror2, error2, x, y As Long
    If Math.Abs(x1I - x2I) < Math.Abs(y1I - y2I) Then
        Utils.Swap x1I, y1I
        Utils.Swap x2I, y2I
        steep = True
    End If
    
    If x1I > x2I Then
        Utils.Swap x1I, x2I
        Utils.Swap y1I, y2I
    End If
    
    dx = x2I - x1I
    dy = y2I - y1I
    derror2 = Math.Abs(dy) * 2
    error2 = 0
    y = y1I
    For x = x1I To x2I
        ' we need to interpolate between the vertices for z
        If steep Then
            PlacePixel (y), (x), 0, v1.VertColor
        Else
            PlacePixel (x), (y), 0, v1.VertColor
        End If
        error2 = error2 + derror2
        If error2 > dx Then
            If y2I > y1I Then
                y = y + 1
            Else
                y = y - 1
            End If
            error2 = error2 - 2 * dx
        End If
    Next x
End Sub

Private Sub SetPixel(ByRef v As Vertex)
    PlacePixel CInt(v.Position.x), CInt(v.Position.y), v.VertColor
End Sub

Private Sub CallVertexShader(ByRef v As Vertex)
    If Not IsEmpty(VertexShaderMethodName) Then
        Application.Run VertexShaderMethodName, v
    End If
End Sub

Private Sub CallFragmentShader()
    If Not IsEmpty(FragmentShaderMethodName) Then
        Application.Run FragmentShaderMethodName, Frag
    End If
End Sub

Private Sub ProcessVertex(ByRef vert As Vertex)
    ' Now that we have the vertex processed
    ' We need to multiply with the projection matrix
    Utils.VectorMatrixMult MVPMatrix, vert.Position
    ' We are in clip space
    ' We need to transform into NDC space
    If vert.Position.w <> 0 Then
        vert.Position.x = vert.Position.x / vert.Position.w
        vert.Position.y = vert.Position.y / vert.Position.w
        vert.Position.z = vert.Position.z / vert.Position.w
        vert.Position.w = 1#
    End If
    ' Pass to vertex shader in NDC coordinate space
    CallVertexShader vert
    
    ' We should have the vertex calculated by now
    ' Go ahead and move it to view screen
    
    vert.Position.x = Int((vert.Position.x * 0.5 + 0.5) * (Surface.width - 1))
    vert.Position.y = Int((vert.Position.y * 0.5 + 0.5) * (Surface.height - 1))
End Sub

Private Sub CalculateModelMatrix()
    ModelTransform.CalculateRotationValues
    
    Dim translateMat As New Matrix
    Utils.IdentMatrix translateMat
    translateMat.MatVal(0, 3) = ModelTransform.Translate.x
    translateMat.MatVal(1, 3) = ModelTransform.Translate.y
    translateMat.MatVal(2, 3) = ModelTransform.Translate.z
    
    Dim scaleMat As New Matrix
    Utils.IdentMatrix scaleMat
    scaleMat.MatVal(0, 0) = ModelTransform.Scaling.x
    scaleMat.MatVal(1, 1) = ModelTransform.Scaling.y
    scaleMat.MatVal(2, 2) = ModelTransform.Scaling.z
    
    Dim rotationMat As New Matrix
    With ModelTransform
        rotationMat.MatVal(0, 0) = .CosRotation.z * .CosRotation.y
        rotationMat.MatVal(0, 1) = -.CosRotation.y * .SinRotation.z
        rotationMat.MatVal(0, 2) = .SinRotation.y
    
        rotationMat.MatVal(1, 0) = .SinRotation.x * .SinRotation.y * .CosRotation.z + .CosRotation.x * .SinRotation.z
        rotationMat.MatVal(1, 1) = -.SinRotation.x * .SinRotation.y * .SinRotation.z + .CosRotation.x * .CosRotation.z
        rotationMat.MatVal(1, 2) = -.SinRotation.x * .CosRotation.y
        
        rotationMat.MatVal(2, 0) = -.CosRotation.x * .SinRotation.y * .CosRotation.z + .SinRotation.x * .SinRotation.z
        rotationMat.MatVal(2, 1) = .CosRotation.x * .SinRotation.y * .SinRotation.z + .SinRotation.x * .CosRotation.z
        rotationMat.MatVal(2, 2) = .CosRotation.x * .CosRotation.y
    
    End With
    
    rotationMat.MatVal(3, 3) = 1
    
    Set Matrices(gMatrixModeEnum.model) = Utils.MatMult3(translateMat, scaleMat, rotationMat).Transpose
    
End Sub

Private Sub DrawTriangle(ByRef v0 As Vertex, ByRef v1 As Vertex, ByRef v2 As Vertex)
    DrawLine v0, v1
    DrawLine v1, v2
    DrawLine v2, v0
End Sub

Private Function EdgeFunctionCW(ByRef v1 As Vec4, ByRef v2 As Vec4, ByRef v3 As Vec4)
    EdgeFunctionCW = (v3.x - v1.x) * (v2.y - v1.y) - (v3.y - v1.y) * (v2.x - v1.x)
End Function

Private Function EdgeFunctionCCW(ByRef v1 As Vec4, ByRef v2 As Vec4, ByRef v3 As Vec4)
    EdgeFunctionCCW = (v1.x - v2.x) * (v3.y - v1.y) - (v1.y - v2.y) * (v3.x - v1.x)
End Function

Private Function TriangleBoundingBox(ByRef v1 As Vec4, ByRef v2 As Vec4, ByRef v3 As Vec4) As BoundingBox
    Dim bb As New BoundingBox
    
    bb.min.x = WorksheetFunction.min(v1.x, v2.x, v3.x)
    bb.min.y = WorksheetFunction.min(v1.y, v2.y, v3.y)
    bb.min.z = WorksheetFunction.min(v1.z, v2.z, v3.z)
    
    bb.max.x = WorksheetFunction.max(v1.x, v2.x, v3.x)
    bb.max.y = WorksheetFunction.max(v1.y, v2.y, v3.y)
    bb.max.z = WorksheetFunction.max(v1.z, v2.z, v3.z)
    
    Set TriangleBoundingBox = bb
End Function

'https://www.scratchapixel.com/lessons/3d-basic-rendering/rasterization-practical-implementation/rasterization-stage
Private Sub RasterTriangle(ByRef v0 As Vertex, ByRef v1 As Vertex, ByRef v2 As Vertex)
    Dim bb As BoundingBox
    Set bb = TriangleBoundingBox(v0.Position, v1.Position, v2.Position)
        
    Dim area As Double
    Dim w0 As Double
    Dim w1 As Double
    Dim w2 As Double
    
    area = EdgeFunctionCW(v0.Position, v1.Position, v2.Position)
    If area < 0.1 Then
        Exit Sub
    End If
    ' temporary vector to calculate the data
    Dim tmpVec As New Vec4
    Dim tmpColor As New Vec4
    
    Dim i As Double
    Dim j As Double
    
    ' We should make them clockwise
       
    v0.VertColor.Divide v0.Position.z
    v1.VertColor.Divide v1.Position.z
    v2.VertColor.Divide v2.Position.z
    
    v0.Position.z = 1# / v0.Position.z
    v1.Position.z = 1# / v1.Position.z
    v2.Position.z = 1# / v2.Position.z
    
    For j = bb.min.y To bb.max.y Step 1
        For i = bb.min.x To bb.max.x Step 1
            tmpVec.x = CDbl(i) + 0.5
            tmpVec.y = CDbl(j) + 0.5
            w0 = EdgeFunctionCW(v1.Position, v2.Position, tmpVec)
            w1 = EdgeFunctionCW(v2.Position, v0.Position, tmpVec)
            w2 = EdgeFunctionCW(v0.Position, v1.Position, tmpVec)
            If w0 >= 0 And w1 >= 0 And w2 >= 0 Then
                w0 = w0 / area
                w1 = w1 / area
                w2 = w2 / area
                
                ' pixel z coordinate
                
                tmpVec.z = 1# / (w0 * v0.Position.z + w1 * v1.Position.z + w2 * v2.Position.z)
                
                ' Color interpolation
                tmpColor.x = tmpVec.z * (w0 * v0.VertColor.x + w1 * v1.VertColor.x + w2 * v2.VertColor.x)
                tmpColor.y = tmpVec.z * (w0 * v0.VertColor.y + w1 * v1.VertColor.y + w2 * v2.VertColor.y)
                tmpColor.z = tmpVec.z * (w0 * v0.VertColor.z + w1 * v1.VertColor.z + w2 * v2.VertColor.z)
                
                
                PlacePixel CDbl(i), CDbl(j), CDbl(tmpVec.z), tmpColor
            End If
        Next i
    Next j
    
End Sub

Private Sub PlacePixel(x As Double, y As Double, z As Double, ByRef vcolor As Vec4)
    Frag.Clear
    Frag.Position.x = x
    Frag.Position.y = y
    Frag.Position.z = z
    Frag.FragColor.Copy vcolor
    Frag.FragColor.Clamp 0, 1
    If x < 0 Or x >= Surface.width Then
        Exit Sub
    End If
    If y < 0 Or y >= Surface.height Then
        Exit Sub
    End If
    If DepthBuffer(CLng(Frag.Position.x) + CLng(Frag.Position.y) * CLng(Surface.width)) < z Then
        Exit Sub
    End If
    
    CallFragmentShader

    DepthBuffer(CLng(Frag.Position.x) + CLng(Frag.Position.y) * CLng(Surface.width)) = Frag.Position.z
    PrimaryBuffer(CLng(Frag.Position.x) + CLng(Frag.Position.y) * CLng(Surface.width)) = RGB(Frag.FragColor.x * 255, Frag.FragColor.y * 255, Frag.FragColor.z * 255)
End Sub

