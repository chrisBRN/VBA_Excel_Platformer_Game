Attribute VB_Name = "windowsAPIs"
Option Explicit

' /////////////////////////////////////////////////////////////////////////////
' WinAPI Functions ////////////////////////////////////////////////////////////
' /////////////////////////////////////////////////////////////////////////////

' https://www.microsoft.com/en-us/download/confirmation.aspx?id=9970

#If Win64 Then
    '64-bit Office
    Public Declare PtrSafe Function GetSystemMetrics Lib "user32.dll" (ByVal index As Long) As Long
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Public Declare PtrSafe Function GetTickCount64 Lib "kernel32" () As LongLong
        
    Public Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
    Public Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
    
    Public Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
        
#Else
    '32-bit Office
    Public Declare Function GetSystemMetrics Lib "user32.dll" (ByVal index As Long) As Long
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Public Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
    
    Public Declare PtrSafe Function SetTimer Lib "user32" (ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
    Public Declare PtrSafe Function KillTimer Lib "user32" (ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr) As Long
    
    Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
        
#End If

' Used with GetSystemMetrics
Public Const SM_CXSCREEN = 0
Public Const SM_CYSCREEN = 1
Public Const SM_CYCAPTION = 4
