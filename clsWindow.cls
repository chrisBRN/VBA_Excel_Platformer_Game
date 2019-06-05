VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private cellsWide As Integer
Private cellsHigh As Integer

Private tLCell As String
Private tRCell As String
Private bLCell As String
Private bRCell As String

Private Sub Class_Initialize()
    
    ActiveWorkbook.Application.DisplayAlerts = False
    
    Call create_sheet
    
    Dim TitleBarHeight As Long
    TitleBarHeight = GetSystemMetrics(SM_CYCAPTION)
    
    ' Sets all cells to be square
    Cells(1, 1).ColumnWidth = 0.5
    With Cells(1, 1)
        Cells.RowHeight = .width
        Cells.ColumnWidth = .ColumnWidth
    End With
            
    ' Set Window dimensions & position
    With Application
        .WindowState = xlNormal ' Avoids error due to prior maximised windowstate
        .Top = 1
        .Left = 1
        .width = Cells(1, 1).height * 128
        .height = (Cells(1, 1).width * 64) - TitleBarHeight + 0.5
        '.Caption = SHEET_NAME 'add games name to titlebar
    End With
    
    ' Assign dimension variables based on window & sheet dimensions
    cellsWide = calc_cells_wide()
    cellsHigh = calc_cells_high()
    
    ' Turns off formula bar, scroll bars, status bar, ribbon, status bar & ribbon
    Call display_features(False)
    
    ' Assign corner cells
    tLCell = Cells(1, 1).Address(1, 1)
    tRCell = Cells(cellsWide, 1).Address(1, 1)
    bLCell = Cells(cellsHigh, 1).Address(1, 1)
    bRCell = Cells(cellsHigh, cellsWide).Address(1, 1)
    
    Cells.Clear ' not sure why this is needed but it seems to help remove latent colours
    
End Sub

Private Sub create_sheet()
    
    ActiveWorkbook.Application.DisplayAlerts = False
    
    'With ThisWorkbook
    
        Dim sh As Worksheet
        Set sh = Worksheets.Add(Before:=Worksheets(1))
        
        'deletes all but first sheet
        Dim i As Integer
        For i = 2 To Worksheets.Count
            Worksheets(i).Delete
        Next i
        
        Worksheets(1).NAME = SHEET_NAME
    
   ' End With
    
    ActiveWorkbook.Application.DisplayAlerts = True

End Sub

Public Sub display_features(display As Boolean)
    
    ' Hides/shows formula bar, scroll bars, status bar & ribbon
    With Application
        .ExecuteExcel4Macro "show.toolbar(""Ribbon""," + CStr(display) + ")"
        .DisplayFormulaBar = display
        .DisplayScrollBars = display
        .DisplayStatusBar = display
    End With
    
    ' Hides/shows gridlines & headings
    ActiveWindow.DisplayGridlines = display
    ActiveWindow.DisplayHeadings = display

End Sub

Private Function calc_cells_wide()

    ' gets the width in "cells" rounded up to the nearest even number
    Dim cellsWide As Integer
    cellsWide = Application.width / Cells(1, 1).width 'truncates to int from double
            
    If (cellsWide Mod 2) <> 0 Then
        cellsWide = cellsWide - 1
    End If
            
    calc_cells_wide = cellsWide
    
End Function

Private Function calc_cells_high()

    ' gets the height in "cells" rounded up to the nearest even number
    Dim cellsHigh As Integer
    cellsHigh = Application.height / Cells(1, 1).height - 1 'truncates to int from double
                      
    If (cellsHigh Mod 2) <> 0 Then
        cellsHigh = cellsHigh - 1
    End If
           
    calc_cells_high = cellsHigh

End Function

' /////////////////////////////////////////////////////////////////////////////
' Getters /////////////////////////////////////////////////////////////////////
' /////////////////////////////////////////////////////////////////////////////

Public Function get_cells_wide()
    get_cells_wide = cellsWide
End Function
Public Function get_cells_high()
    get_cells_high = cellsHigh
End Function
Public Function get_top_left_cell()
    get_top_left_cell = tLCell
End Function
Public Function get_top_right_cell()
    get_top_right_cell = tRCell
End Function
Public Function get_bottom_left_cell()
    get_bottom_left_cell = bLCell
End Function
Public Function get_bottom_right_cell()
    get_bottom_right_cell = bRCell
End Function
Public Function get_sheet() As Worksheet
    Set get_sheet = Worksheets(SHEET_NAME)
End Function


