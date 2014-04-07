Attribute VB_Name = "modTimer"
Option Explicit


Private oTimer   As clsTimer
Private lTimerID As Long

Declare Function SetTimer Lib "user32" _
      (ByVal hWnd As Long, _
      ByVal nIDEvent As Long, _
      ByVal uElapse As Long, _
      ByVal lpTimerFunc As Long) As Long

Declare Function KillTimer Lib "user32" _
      (ByVal hWnd As Long, _
      ByVal nIDEvent As Long) As Long

Public Sub TimerCallBack(ByVal hWnd As Long, _
                          ByVal uMsg As Long, _
                          ByVal idEvent As Long, _
                          ByVal dwTime As Long)

   oTimer.RaiseTimerEvent

End Sub

Public Function TimerStart(ByRef oTmr As clsTimer, _
                           lInterval As Long, ByVal hWnd As Long) As Long
    Set oTimer = oTmr
    lTimerID = SetTimer(hWnd, 0, lInterval, AddressOf TimerCallBack)
    TimerStart = lTimerID

End Function

Public Function TimerStop(ByVal hWnd As Long) As Long
    TimerStop = KillTimer(0, lTimerID)
    Set oTimer = Nothing
End Function



