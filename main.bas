Attribute VB_Name = "main"
Option Explicit
  
Public window As clsWindow
Public lv As clsLevel
Public player As clsPlayer

Public IsRunning As Boolean ' TODO Replace this with a more detailed state class

Public Const SHEET_NAME = "Important Work"
Public Const PLAYER_COLOUR = rgbCrimson
       
Sub main()
    
    ' Set up
    Application.ScreenUpdating = True
    Application.Interactive = False
    Application.Calculation = xlCalculationManual
    Call block_keyboard_keys
    Set window = New clsWindow
    Set lv = New clsLevel
    Set player = New clsPlayer
    
    ' Timing set up
    Dim startTime As Long
    startTime = 0
    Dim FPS As Integer
    FPS = (1000 / 50)
    
    ' Main Game Loop
    IsRunning = True
    While IsRunning
    
        startTime = GetTickCount
    
        Call update
        Call draw
        
        Sleep (FPS - (startTime - GetTickCount))
    
    Wend
    
    ' Clean Up
    player.get_curCell.Clear
    Application.Calculation = xlSemiautomatic
    Application.ScreenUpdating = True
        
End Sub

Sub update()
    player.update
End Sub
Sub draw()
    player.draw
End Sub

Private Sub block_keyboard_keys() 'TODO Turn these back on on close
    Application.OnKey "{ESC}", ""
    Application.OnKey "{UP}", ""
    Application.OnKey "{DOWN}", ""
    Application.OnKey "{LEFT}", ""
    Application.OnKey "{RIGHT}", ""
    Application.OnKey "{TAB}", ""
End Sub
   
Public Sub total_exit_game()
    window.display_features (True)
    Call exit_current_game
End Sub

Public Sub exit_current_game()
    IsRunning = False
    Application.Interactive = True
    Application.ScreenUpdating = True
End Sub
