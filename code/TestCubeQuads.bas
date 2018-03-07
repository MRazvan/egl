Attribute VB_Name = "TestCubeQuads"
Public ModelRotation As Double
Private Disp As Display

Public Sub Initialize(testPsz As Integer, testW As Integer, testH As Integer, testShName As String)
    Set Disp = New Display
    Disp.Initialize testPsz, testW, testH, testShName
    'Disp.AddBackBuffer "Buffer2"
    
    EGL.gInitialize Disp
    
    EGL.gSetVertexShader "SimpleShader.VertexShader"
    EGL.gSetFragmentShader "SimpleShader.FragmentShader"
    
    EGL.gMatrixMode gMatrixModeEnum.projection
    EGLU.gluPerspective 45, CDbl(Disp.width) / CDbl(Disp.height), 0.1, 10000
    
    EGL.gMatrixMode gMatrixModeEnum.model
    
    EGL.gClearColor 0, 0, 0
    EGL.gTranslate 0, 0, -4
    
    EGL.gColor3b 255, 0, 0
    
    ModelRotation = 0
End Sub

Public Sub Run()
        EGL.gRotate 45, ModelRotation, 0
        EGL.gClear ColorBit
        EGL.gBegin Quads
                ' Border around the screen
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
        Run
    Next loopCnt
    SecondsElapsed = Round(Timer - StartTime, 2)
    
    Dim fps As Double
    fps = Round((loopMax) / SecondsElapsed, 2)
    
    Dim pixelFillRate As Double
    pixelFillRate = Round((CDbl(loopMax) * CDbl(Disp.width) * CDbl(Disp.height)) / SecondsElapsed, 2)
    Application.Calculation = xlCalculationAutomatic
    MsgBox "This code ran successfully in " & SecondsElapsed & " seconds" & Chr(13) & Chr(10) & "FPS : " & fps & Chr(13) & Chr(10) & "Pixel Fill rate : " & pixelFillRate, vbInformation
End Sub
