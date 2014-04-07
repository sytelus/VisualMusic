VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event Timer()
Private blnEnabled As Boolean
Private lTimerID   As Long
Private lInterval  As Long
Public hWnd As Long

Public Property Let Interval(lVal As Long)
    lInterval = lVal
End Property
Public Property Get Interval() As Long
    Interval = lInterval
End Property
Public Property Let Enabled(blnVal As Boolean)
    If blnVal = False Then
        StopTimer
    Else
        StartTimer
    End If
End Property

Private Function StartTimer() As Boolean

    If Not blnEnabled Then
        lTimerID = TimerStart(Me, lInterval, hWnd)
        If lTimerID = 0 Then
          err.Raise 60000 + vbObjectError, "clsTimer", "Could not start Timer"
        End If
        blnEnabled = True
    End If
    StartTimer = True

End Function

Private Function StopTimer() As Boolean

    If blnEnabled Then
        lTimerID = TimerStop(hWnd)
        If lTimerID = 0 Then
          err.Raise 60001 + vbObjectError, "clsTimer", "Could not stop Timer"
        End If
        blnEnabled = False
    End If
    StopTimer = False

End Function

Friend Function RaiseTimerEvent()

   RaiseEvent Timer

End Function

Private Sub Class_Initialize()
    hWnd = 0
End Sub

Private Sub Class_Terminate()
    Call TimerStop(hWnd)
End Sub

