Attribute VB_Name = "main"
Option Explicit

' /////////////////////////////////////////////////////////////////////////////
' GLOBALS /////////////////////////////////////////////////////////////////////
' /////////////////////////////////////////////////////////////////////////////

Private GameTimerID As Long
Private startTime As Long

Public Const SHEET_NAME = "Important Work"

Public window As clsWindow
Public lv As clsLevel
Public player As clsPlayer

Public Const GRAVITY As Byte = 1

' /////////////////////////////////////////////////////////////////////////////
' PROGRAM /////////////////////////////////////////////////////////////////////
' /////////////////////////////////////////////////////////////////////////////
    
Sub main()
    
    
    
    
    Call TerminateTimer

    block_keyboard_keys

    Application.DisplayAlerts = False

    Set window = New clsWindow
    Set lv = New clsLevel
    Set player = New clsPlayer
    
    ThisWorkbook.Activate

    InitialiseTimer ' runs game
    
    
    
End Sub

Public Sub InitialiseTimer()

    'the time in milliseconds between each tick
    Dim GameTimerInterval As Double
    
    'makes the game clock tick 20 times per second
    GameTimerInterval = 30
    
    'pauses for half a second before starting game
    Sleep (500)
    
    'starts the timer calling UpdateGame
    GameTimerID = SetTimer(0, 0, GameTimerInterval, AddressOf UpdateAndDrawGame)

End Sub

Public Sub TerminateTimer()

    If GameTimerID <> 0 Then
        'stops the timer whose ID we stored earlier
        KillTimer 0, GameTimerID
        GameTimerID = 0
    End If

End Sub

Sub UpdateAndDrawGame()
    
    Application.ScreenUpdating = False
    
    player.update
    player.draw
    
    Application.ScreenUpdating = True
    
End Sub

Private Sub block_keyboard_keys()
'TODO add opposite into game exit/close functions

    Application.OnKey "{ESC}", ""
    Application.OnKey "{UP}", ""
    Application.OnKey "{DOWN}", ""
    Application.OnKey "{LEFT}", ""
    Application.OnKey "{RIGHT}", ""
    Application.OnKey "{TAB}", ""

End Sub

Public Sub total_exit_game()

    Sleep (500)

    Application.ScreenUpdating = False
    Call TerminateTimer
    window.display_features (True)
    Application.Interactive = True
    Application.ScreenUpdating = True
    
End Sub

Public Sub exit_current_game()

    Call TerminateTimer
   
End Sub
