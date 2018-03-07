VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Display"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public width As Integer
Public height As Integer

Private PixelSz As Integer
Private SheetName As String

Private Buffers() As SheetBuffer
Private BuffersCount As Integer
Private CurrentBufferIdx As Integer

Public Sub Fill(r As Integer, g As Integer, b As Integer)
    Buffers(CurrentBufferIdx).TargetSheet.Activate
    BackgroundColor = RGB(r, g, b)
    
    ActiveWindow.GridlineColor = RGB(r, g, b)
    Buffers(CurrentBufferIdx).DisplayRange.Interior.color = RGB(r, g, b)
End Sub

Public Sub LockScreen(doLock As Boolean)
    Application.ScreenUpdating = Not doLock
End Sub

Public Sub AddBackBuffer(bufferName As String)
    LockScreen False
    
    BuffersCount = BuffersCount + 1
    ReDim Preserve Buffers(BuffersCount)
   
    Dim buf As New SheetBuffer
    buf.Initialize bufferName, width, height, PixelSz
    Set Buffers(BuffersCount - 1) = buf
    
    DoEvents
    LockScreen True
End Sub

Public Sub SwapBuffers()
    Buffers(CurrentBufferIdx).TargetSheet.Activate
    CurrentBufferIdx = CurrentBufferIdx + 1
    If CurrentBufferIdx >= BuffersCount Then
        CurrentBufferIdx = 0
    End If
    If BuffersCount > 1 Then
        Buffers(CurrentBufferIdx).Clear
    End If
End Sub


Public Sub DisplayBuffer(ByRef buf() As Long)
    Dim idx As Long
    idx = 0
    Dim pixels As Range
    Set pixels = Buffers(CurrentBufferIdx).DisplayRange.Cells
    LockScreen True
    For Each p In pixels
        p.Interior.color = buf(idx)
        idx = idx + 1
    Next p
    LockScreen False
End Sub



Private Sub InitializeDisplay()
    HasBackBuffer = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayStatusBar = False
    
    BuffersCount = 0
    CurrentBufferIdx = 0

    AddBackBuffer SheetName
    DropUnusedStyles
    
    DoEvents
    LockScreen True
    
End Sub

Public Sub Initialize(pixelSize As Integer, screenWidth As Integer, screenHeight As Integer, targetSheetName As String)

    PixelSz = pixelSize
    width = screenWidth
    height = screenHeight
    SheetName = targetSheetName

    InitializeDisplay
End Sub

Private Sub DropUnusedStyles()
    Dim styleObj As Style
    Dim rngCell As Range
    Dim wb As Workbook
    Dim wsh As Worksheet
    Dim str As String
    Dim iStyleCount As Long
    Dim dict As New Scripting.Dictionary    ' <- from Tools / References... / "Microsoft Scripting Runtime"
    
    ' wb := workbook of interest.  Choose one of the following
    ' Set wb = ThisWorkbook ' choose this module's workbook
    Set wb = ActiveWorkbook ' the active workbook in excel
    Debug.Print "BEGINNING # of styles in workbook: " & wb.Styles.Count
    ' dict := list of styles
    For Each styleObj In wb.Styles
        str = styleObj.NameLocal
        iStyleCount = iStyleCount + 1
        Call dict.Add(str, 0)    ' First time:  adds keys
    Next styleObj
    Debug.Print "  dictionary now has " & dict.Count & " entries."
    ' Status, dictionary has styles (key) which are known to workbook


    ' Traverse each visible worksheet and increment count each style occurrence
    For Each wsh In wb.Worksheets
        If wsh.Visible Then
            For Each rngCell In wsh.UsedRange.Cells
                str = rngCell.Style
                dict.Item(str) = dict.Item(str) + 1     ' This time:  counts occurrences
            Next rngCell
        End If
    Next wsh
    ' Status, dictionary styles (key) has cell occurrence count (item)


    ' Try to delete unused styles
    Dim aKey As Variant
    On Error Resume Next    ' wb.Styles(aKey).Delete may throw error

    For Each aKey In dict.Keys
        If dict.Item(aKey) = 0 Then
            ' Occurrence count (Item) indicates this style is not used
            Call wb.Styles(aKey).Delete
            If err.Number <> 0 Then
                Debug.Print vbTab & "^-- failed to delete"
                err.Clear
            End If
            Call dict.Remove(aKey)
        End If
    Next aKey
End Sub
