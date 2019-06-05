VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const GROUND_COLOUR1 = rgbBurlyWood
Private Const GROUND_COLOUR2 = rgbBisque
Private Const GRASS1 = rgbLightGreen
Private Const GRASS2 = rgbLightSeaGreen

Private ground As Range
Private floor As Range
Private roof As Range
Private leftEdge As Range
Private rightEdge As Range

Private Sub Class_Initialize()
    
    Dim bLCell As String
    Dim bRCell As String
    
    bLCell = window.get_bottom_left_cell
    bRCell = window.get_bottom_right_cell
    
    Dim groundThickness As Integer
    
    groundThickness = window.get_cells_high / 7
    
    With Worksheets(SHEET_NAME)
        Set ground = Range(bLCell, bRCell).Offset(-groundThickness, 0).Resize(groundThickness)
    End With
    
    ground.Interior.Color = GROUND_COLOUR1
    Call dither(ground, GROUND_COLOUR2)
    Call add_top_edge_colour(ground, GRASS1, GRASS2)
    
End Sub

Private Sub dither(targetRange As Range, OVERLAY_COLOUR)
    Dim cell As Range
    For Each cell In targetRange
        If cell.Row Mod 2 = 0 And cell.Column Mod 2 <> 0 Then
            cell.Interior.Color = OVERLAY_COLOUR
        ElseIf cell.Row Mod 2 <> 0 And cell.Column Mod 2 = 0 Then
            cell.Interior.Color = OVERLAY_COLOUR
        End If
    Next cell
End Sub

Private Sub add_top_edge_colour(targetRange As Range, MAIN_COLOUR, ACCENT_COLOUR)
    With targetRange.Rows(1).Offset(-1, 0).Interior
        .Color = MAIN_COLOUR
        .Pattern = xlLightUp
        .PatternColor = ACCENT_COLOUR
    End With
End Sub

Public Function get_ground_range() As Range
    Set get_ground_range = ground
End Function
Public Function get_floor() As Range
    Set get_floor = ground.Rows(1).Offset(-1, 0)
End Function
Public Function get_roof() As Range
    'Set get_ground_range = ground
End Function
Public Function get_left_edge() As Range
    'Set get_ground_range = ground
End Function

