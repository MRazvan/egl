Attribute VB_Name = "TestRenderer"
' Issues
'   The y coordinate is inverted ... need to figure out why
'   The mesh drawing (Line drawing) is not taking into account the depthbuffer, need to find a different way to draw the lines that takes into account the z buffering
'   Excel has a limit on the number of formats it can support (in our case colors in the cells)
'   We only have support for one model matrix, maybe we should implement a newer version of GL and let the user handle the matrices with the vertex shader. Or implement a stack of matrices

Public ModelRotation As Vec3
Private Disp As Display
Private StartTime As Double
Private SecondsElapsed As Double

Public Sub Initialize(testPsz As Integer, testW As Integer, testH As Integer, testShName As String)
    Set ModelRotation = New Vec3
    ModelRotation.Clear
    Set Disp = New Display
    Disp.Initialize testPsz, testW, testH, testShName
    
    EGL.gInitialize Disp
    EGL.gClearColor 0, 0, 0
    'Shaders
    EGL.gSetVertexShader "SimpleShader.VertexShader"
    EGL.gSetFragmentShader "SimpleShader.FragmentShader"
    
    ' Matrices
    EGL.gMatrixMode gMatrixModeEnum.projection
    EGLU.gluPerspective 45, CDbl(Disp.width) / CDbl(Disp.height), 0.1, 10000
    
    EGL.gMatrixMode gMatrixModeEnum.model
    
    ' Translate the model
    EGL.gTranslate 0, 0, -4
End Sub

Public Sub RunQuads()
        EGL.gRotate ModelRotation.x, ModelRotation.y, ModelRotation.z
        EGL.gClear ColorBit Or DepthBit
        EGL.gBegin Quads
            EGL.gColor3d 1, 0, 0
            EGL.gVertex3d 1#, 1#, 1#
            EGL.gVertex3d 1#, -1#, 1#
            EGL.gVertex3d 1#, -1#, -1#
            EGL.gVertex3d 1#, 1#, -1#
            EGL.gVertex3d 1#, 1#, -1#
            EGL.gVertex3d 1#, -1#, -1#
            EGL.gVertex3d -1#, -1#, -1#
            EGL.gVertex3d -1#, 1#, -1#
            EGL.gVertex3d -1#, 1#, -1#
            EGL.gVertex3d -1#, -1#, -1#
            EGL.gVertex3d -1#, -1#, 1#
            EGL.gVertex3d -1#, 1#, 1#
            EGL.gVertex3d -1#, 1#, 1#
            EGL.gVertex3d -1#, -1#, 1#
            EGL.gVertex3d 1#, -1#, 1#
            EGL.gVertex3d 1#, 1#, 1#
            EGL.gVertex3d -1#, 1#, -1#
            EGL.gVertex3d -1#, 1#, 1#
            EGL.gVertex3d 1#, 1#, 1#
            EGL.gVertex3d 1#, 1#, -1#
            EGL.gVertex3d -1#, -1#, 1#
            EGL.gVertex3d -1#, -1#, -1#
            EGL.gVertex3d 1#, -1#, -1#
            EGL.gVertex3d 1#, -1#, 1#
        EGL.gEnd
        
        EGL.gFlush
End Sub

Public Sub RunTriangles(dofill As Boolean)
        If dofill Then
            EGL.gPolygonMode gPolygonModeEnum.Fill
        Else
            EGL.gPolygonMode gPolygonModeEnum.Line
        End If
        EGL.gRotate ModelRotation.x, ModelRotation.y, ModelRotation.z
        EGL.gClear ColorBit Or DepthBit
        EGL.gBegin Triangle
            ' Front
            EGL.gColor3d 1, 0, 0
            EGL.gVertex3d -1, 1, 1
            
            EGL.gColor3d 0, 0, 1
            EGL.gVertex3d 1, -1, 1
            EGL.gColor3d 1, 1, 0
            
            EGL.gVertex3d -1, -1, 1
            EGL.gColor3d 0, 0, 1
            EGL.gVertex3d 1, -1, 1
            
            EGL.gColor3d 1, 0, 0
            EGL.gVertex3d -1, 1, 1
            EGL.gColor3d 0, 1, 0
            EGL.gVertex3d 1, 1, 1
            
            'Right
            EGL.gColor3d 0, 1, 0
            EGL.gVertex3d 1, 1, 1
            EGL.gVertex3d 1, 1, -1
            EGL.gVertex3d 1, -1, 1
            EGL.gVertex3d 1, -1, 1
            EGL.gVertex3d 1, 1, -1
            EGL.gVertex3d 1, -1, -1
            
            'Back
            EGL.gColor3d 0, 0, 1
            EGL.gVertex3d 1, -1, -1
            EGL.gVertex3d 1, 1, -1
            EGL.gVertex3d -1, 1, -1
            EGL.gVertex3d 1, -1, -1
            EGL.gVertex3d -1, 1, -1
            EGL.gVertex3d -1, -1, -1
            
            'Right
            EGL.gColor3d 1, 0, 1
            EGL.gVertex3d -1, -1, 1
            EGL.gVertex3d -1, 1, -1
            EGL.gVertex3d -1, 1, 1
            EGL.gVertex3d -1, -1, 1
            EGL.gVertex3d -1, -1, -1
            EGL.gVertex3d -1, 1, -1
            
            'Bottom
            EGL.gColor3d 0, 1, 1
            EGL.gVertex3d -1, 1, 1
            EGL.gVertex3d 1, 1, -1
            EGL.gVertex3d 1, 1, 1
            EGL.gVertex3d -1, 1, 1
            EGL.gVertex3d -1, 1, -1
            EGL.gVertex3d 1, 1, -1
            
            'Top
            EGL.gColor3d 1, 1, 0
            EGL.gVertex3d 1, -1, 1
            EGL.gVertex3d 1, -1, -1
            EGL.gVertex3d -1, -1, -1
            EGL.gVertex3d -1, -1, -1
            EGL.gVertex3d -1, -1, 1
            EGL.gVertex3d 1, -1, 1
            
        EGL.gEnd
        
        EGL.gFlush
End Sub

Public Sub RunSingleTriangle()
        EGL.gPolygonMode gPolygonModeEnum.Fill
        EGL.gRotate ModelRotation.x, ModelRotation.y, ModelRotation.z
        EGL.gClear ColorBit Or DepthBit
        EGL.gBegin Triangle
            EGL.gColor3d 0, 0, 1
            EGL.gVertex3d -1, 1, 1
            EGL.gColor3d 0, 1, 0
            EGL.gVertex3d 1, 1, 1
            EGL.gColor3d 1, 0, 0
            EGL.gVertex3d 0, -1, 1
        EGL.gEnd
        EGL.gFlush
End Sub

Public Sub RunTriangle()
        EGL.gPolygonMode gPolygonModeEnum.Line
        EGL.gRotate ModelRotation.x, ModelRotation.y, ModelRotation.z
        EGL.gClear ColorBit Or DepthBit
        EGL.gColor3d 1, 0, 0
        EGL.gBegin Triangle
            EGL.gVertex3d -1, 1, 1
            EGL.gVertex3d 1, 1, 1
            EGL.gVertex3d 0, -1, 1
        EGL.gEnd
        EGL.gFlush
End Sub

Public Sub PerfTest()
    Dim Disp As New Display
    Disp.Initialize 4, 100, 75, "Display"
    EGL.gInitialize Disp
    
    EGL.gSetVertexShader "SimpleShader.VertexShader"
    EGL.gSetFragmentShader "SimpleShader.FragmentShader"
    
    EGL.gMatrixMode gMatrixModeEnum.projection
    EGLU.gluPerspective 45, CDbl(Disp.width) / CDbl(Disp.height), 0.1, 10000
    
    EGL.gMatrixMode gMatrixModeEnum.model
    
    EGL.gClearColor 0, 0, 0
    EGL.gTranslate 0, 0, -4

    EGL.gColor3b 255, 0, 0
    
    Dim loopCnt, loopMax As Integer
    loopMax = 10
    StartTime = Timer
    For loopCnt = 1 To loopMax
        RunTriangles
    Next loopCnt
    SecondsElapsed = Round(Timer - StartTime, 2)
    
    Dim fps As Double
    fps = Round((loopMax) / SecondsElapsed, 2)
    
    Dim pixelFillRate As Double
    pixelFillRate = Round((CDbl(loopMax) * CDbl(Disp.width) * CDbl(Disp.height)) / SecondsElapsed, 2)
    Application.Calculation = xlCalculationAutomatic
    MsgBox "This code ran successfully in " & SecondsElapsed & " seconds" & Chr(13) & Chr(10) & "FPS : " & fps & Chr(13) & Chr(10) & "Pixel Fill rate : " & pixelFillRate, vbInformation
End Sub
