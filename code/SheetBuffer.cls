VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SheetBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public TargetSheet As Worksheet
Public DisplayRange As Range
Public SheetName As String

Public Sub Initialize(name As String, w As Integer, h As Integer, PixelSz As Integer)
    SheetName = name
    Set TargetSheet = Worksheets(SheetName)
    TargetSheet.Activate
    TargetSheet.DisplayPageBreaks = False
        ' Disable selections
    With TargetSheet
     .EnableSelection = xlNoSelection
     .Protect Contents:=False, UserInterfaceOnly:=False
    End With
    
    With ActiveWindow
        .Zoom = 100
        .DisplayHeadings = False
        .DisplayFormulas = False
        .DisplayGridlines = False
        .DisplayRuler = False
        .DisplayHorizontalScrollBar = False
        .DisplayVerticalScrollBar = False
        .DisplayOutline = False
    End With

    TargetSheet.Range(TargetSheet.Cells(1, 1), TargetSheet.Cells(1000, 1000)).Clear
    DoEvents
    
    ' Get the range that coresponds to the display resolution
    With TargetSheet
        ' Hide everything, we will show it after initialization
        .Columns.Hidden = True
        .Rows.Hidden = True
        Set DisplayRange = .Range(.Cells(1, 1), .Cells(h, w))
        ' Set the size and style of the cells
        DisplayRange.RowHeight = 1 * PixelSz
        DisplayRange.ColumnWidth = 0.1 * PixelSz
        For Each Border In DisplayRange.Borders
            Border.LineStyle = Excel.XlLineStyle.xlLineStyleNone
        Next
        DisplayRange.EntireRow.Hidden = False
        DisplayRange.EntireColumn.Hidden = False
        DisplayRange.Select
    End With
    TargetSheet.ScrollArea = "A1:A1"

    ' Disable selections
    With TargetSheet
     .EnableSelection = xlNoSelection
     .EnableCalculation = False
     .EnableFormatConditionsCalculation = False
     .Protect Contents:=True, UserInterfaceOnly:=True
    End With
    
End Sub

Public Sub Clear()
    DisplayRange.ClearFormats
End Sub
