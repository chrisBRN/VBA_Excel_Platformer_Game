VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const PLAYER_COLOUR = rgbCrimson

Private preCell As Range
Private curCell As Range

Private xMove As Long
Private yMove As Long
Private floor As Range
Private roof As Range


Private preUpKeyState As Integer
Private preDownKeyState As Integer
Private preLeftKeyState As Integer
Private preRightKeyState As Integer

Private Sub Class_Initialize()

    preUpKeyState = GetAsyncKeyState(vbKeyUp)
    preDownKeyState = GetAsyncKeyState(vbKeyDown)
    preLeftKeyState = GetAsyncKeyState(vbKeyLeft)
    preRightKeyState = GetAsyncKeyState(vbKeyRight)
    
    Set curCell = Range("H12")
    Set preCell = curCell
    xMove = 0
    yMove = 0
    
    curCell.Interior.Color = PLAYER_COLOUR
    
    ' likely need to move into collison function
    Set floor = lv.get_ground_range().Rows(1).Offset(-2, 0)
    
End Sub

Private Sub check_keyboard()
    
' TODO remove the function calls for movement here if they aren't used elsewhere (to avoid needless function calls)
    
    If GetAsyncKeyState(vbKeyUp) <> 0 And preUpKeyState = 0 Then move_up
    If GetAsyncKeyState(vbKeyDown) <> 0 And preDownKeyState = 0 Then move_down
    If GetAsyncKeyState(vbKeyLeft) <> 0 Then move_left
    If GetAsyncKeyState(vbKeyRight) <> 0 Then move_right
    If GetAsyncKeyState(vbKeyEscape) <> 0 Then total_exit_game
    If GetAsyncKeyState(vbKeyTab) <> 0 Then exit_current_game
    
    preUpKeyState = GetAsyncKeyState(vbKeyUp)
    preDownKeyState = GetAsyncKeyState(vbKeyDown)

End Sub

Private Sub move_up()
    xMove = xMove - 12
End Sub
Private Sub move_down()
    xMove = xMove + 1
End Sub
Private Sub move_left()
    yMove = yMove - 1
End Sub
Private Sub move_right()
    yMove = yMove + 1
End Sub


Public Sub update()

    Dim targetRow As Integer
    Dim targetCol As Integer

    check_keyboard

    xMove = xMove + GRAVITY
    Set preCell = curCell ' before
    
    targetRow = preCell.Row + xMove
    targetCol = preCell.Column + yMove
    
    ' Bounds checking
    If targetRow >= floor.Row Then targetRow = floor.Row
    If targetRow <= 1 Then targetRow = 1
    If targetCol >= floor.Columns.Count Then targetCol = floor.Columns.Count
    If targetCol <= 1 Then targetCol = 1
   
    Set curCell = Cells(targetRow, targetCol)
    yMove = 0
    xMove = 0 + GRAVITY

End Sub

Public Sub draw()
    
    preCell.Clear
    curCell.Interior.Color = PLAYER_COLOUR

End Sub





