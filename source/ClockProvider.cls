VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClockProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long

Public Event ClockPulse()

Public Sub GenerateClockPulse()
    RaiseEvent ClockPulse
End Sub

Public Sub Start()
    Dim lStartTime As Long
    lStartTime = GetTickCount
    Do
        DoEvents
        Do While GetTickCount - lStartTime > 50
            DoEvents
        Loop
        RaiseEvent ClockPulse
    Loop While True
End Sub
