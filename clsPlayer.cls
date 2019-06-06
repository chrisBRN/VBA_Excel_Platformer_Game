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

Private Const PLAYER_COLOUR2 = rgbCrimson

Private preCell As Range
Private curCell As Range

Private xMove As Double
Private yMove As Double
Private floor As Range
Private roof As Range

Private isMoving As Boolean

Private preUpKeyState As Integer
Private preDownKeyState As Integer
Private preLeftKeyState As Integer
Private preRightKeyState As Integer

Private gravity As Byte

Private Sub Class_Initialize()

    preUpKeyState = GetAsyncKeyState(vbKeyUp)
    preDownKeyState = GetAsyncKeyState(vbKeyDown)
    preLeftKeyState = GetAsyncKeyState(vbKeyLeft)
    preRightKeyState = GetAsyncKeyState(vbKeyRight)
    
    Set curCell = Range("H12")
    Set preCell = curCell
    xMove = 0
    yMove = 0
    gravity = 1
    
    curCell.Interior.Color = PLAYER_COLOUR
    
    ' TODO Build full collision system
    Set floor = lv.get_ground_range().Rows(1).Offset(-2, 0)
    
End Sub

Private Sub check_keyboard()
    
    ' Checks for keyboard input from the player

    ' Movement controls
    If GetAsyncKeyState(vbKeyUp) <> 0 And preUpKeyState = 0 Then move_up
    If GetAsyncKeyState(vbKeyDown) <> 0 And preDownKeyState = 0 Then move_down
    If GetAsyncKeyState(vbKeyLeft) <> 0 Then move_left
    If GetAsyncKeyState(vbKeyRight) <> 0 Then move_right
    
    ' Stores prior key state to avoid continuous input and the associated function calls
    preUpKeyState = GetAsyncKeyState(vbKeyUp)
    preDownKeyState = GetAsyncKeyState(vbKeyDown)
    
    ' Exiting game
    If GetAsyncKeyState(vbKeyEscape) <> 0 Then total_exit_game
    If GetAsyncKeyState(vbKeyTab) <> 0 Then exit_current_game
    
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
    
    ' Clean up
    isMoving = False
    gravity = 1

    ' Gets player input
    check_keyboard
    
    ' Stores the players next cell location
    ' Allows for pre-move bounds checking
    Dim targetRow As Integer
    Dim targetCol As Integer
    xMove = xMove + gravity
    Set preCell = curCell
    targetRow = preCell.Row + xMove
    targetCol = preCell.Column + yMove
    
    ' Bounds checking
    If targetRow >= floor.Row Then
        targetRow = floor.Row
        xMove = xMove - gravity
    End If
    If targetRow <= 1 Then targetRow = 1
    If targetCol >= floor.Columns.Count Then targetCol = floor.Columns.Count
    If targetCol <= 1 Then targetCol = 1
   
    ' Update player cell location
    Set curCell = Cells(targetRow, targetCol)
    
    ' Allows a reduction draw calls (when the player isn't moving)
    If xMove <> 0 Or yMove <> 0 Then isMoving = True
        
    ' Reset to prevent exponential movement
    yMove = 0
    xMove = 0 + gravity

End Sub

Public Sub draw()
    
    ' Clears old player location and 'draws' the new one
    ' TODO This should be done with a buffer
    ' TODO Old player location should revert to it's prior "graphic" rather than clear
    If isMoving Then
        Application.ScreenUpdating = True
        preCell.Clear
        curCell.Interior.Color = PLAYER_COLOUR
        Application.ScreenUpdating = False
    End If
    
End Sub
Public Function get_preCell()
    Set get_preCell = preCell
End Function
Public Function get_curCell()
    Set get_curCell = curCell
End Function
Public Function get_isMoving()
    Set get_isMoving = isMoving
End Function
