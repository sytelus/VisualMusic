Attribute VB_Name = "modMIDIEngine"
Option Explicit

Private Const MMSYSERR_NOERROR = 0  '  no error
Private Const MMSYSERR_BASE = 0
Private Const MMSYSERR_BADERRNUM = (MMSYSERR_BASE + 9)
Private Const MMSYSERR_INVALPARAM = (MMSYSERR_BASE + 11)

Private Declare Function midiOutClose Lib "winmm.dll" (ByVal hMidiOut As Long) As Long
Private Declare Function midiOutOpen Lib "winmm.dll" (lphMidiOut As Long, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Private Declare Function midiOutGetErrorText Lib "winmm.dll" Alias "midiOutGetErrorTextA" (ByVal err As Long, ByVal lpText As String, ByVal uSize As Long) As Long

Private mlMIDIOutHandle As Long
Private mlMIDIOutHandleUserCount As Long

Private mbaChannelFreeStatus(1 To 15) As Boolean

Public Enum MIDIEngineProperties
    mepInstrument = 0
    mepVolume = 1
    mepOctave = 2
    mepPan = 3
End Enum

Public Const lMIDI_ENG_DEFAULT_INSTRUMENT As Long = 25
Public Const lMIDI_ENG_DEFAULT_OCTAVE As Long = 4
Public Const lMIDI_ENG_DEFAULT_PAN As Long = 64
Public Const lMIDI_ENG_DEFAULT_VOLUME As Long = 100


Public Function GetMIDIOutHandle() As Long
    
    Dim lMidiAPIReturn As Long
    
    If mlMIDIOutHandle = 0 Then
        
        lMidiAPIReturn = midiOutOpen(mlMIDIOutHandle, -1, 0, 0, 0)
        If mlMIDIOutHandle <> 0 Then
            Call SaveSetting("Visual Music", "Last Values", "MIDIHandle", mlMIDIOutHandle)
        End If
        
        If lMidiAPIReturn <> MMSYSERR_NOERROR Then
            Dim lLastHandleValue As Long
            lLastHandleValue = GetSetting("Visual Music", "Last Values", "MIDIHandle", "0")
            Call midiOutClose(lLastHandleValue)
            lMidiAPIReturn = midiOutOpen(mlMIDIOutHandle, -1, 0, 0, 0)
            If mlMIDIOutHandle <> 0 Then
                Call SaveSetting("Visual Music", "Last Values", "MIDIHandle", mlMIDIOutHandle)
            End If
        End If
        
        Call CheckMidiApiReturn(lMidiAPIReturn)
        
        Dim lChannelFreeStatusArrayIndex As Long
        
        For lChannelFreeStatusArrayIndex = LBound(mbaChannelFreeStatus) To UBound(mbaChannelFreeStatus)
            mbaChannelFreeStatus(lChannelFreeStatusArrayIndex) = True
        Next lChannelFreeStatusArrayIndex
    
    End If
    
    mlMIDIOutHandleUserCount = mlMIDIOutHandleUserCount + 1
    
    GetMIDIOutHandle = mlMIDIOutHandle
    
End Function

Public Sub ReleaseMIDIOutHandle()
    
    mlMIDIOutHandleUserCount = mlMIDIOutHandleUserCount - 1
    
    If mlMIDIOutHandleUserCount = 0 Then
        If mlMIDIOutHandle <> 0 Then
            Call midiOutClose(mlMIDIOutHandle)
        End If
    End If
        
End Sub

Public Sub CheckMidiApiReturn(ByVal vlMidiAPIReturn As Long)
    If vlMidiAPIReturn <> MMSYSERR_NOERROR Then
        Dim sErrorText As String
        Dim lGetErrorMessageAPIReturn As Long
        Const lERROR_MESSAGE_LEN As Long = 255
        sErrorText = String$(lERROR_MESSAGE_LEN, " ")
        lGetErrorMessageAPIReturn = midiOutGetErrorText(vlMidiAPIReturn, sErrorText, lERROR_MESSAGE_LEN)
        If lGetErrorMessageAPIReturn <> MMSYSERR_BADERRNUM Or lGetErrorMessageAPIReturn <> MMSYSERR_INVALPARAM Then
            sErrorText = Trim$(sErrorText)
        Else
            sErrorText = "MIDI API failed. Error information not available."
        End If
        err.Raise lERR_BASE_MIDIENG + vlMidiAPIReturn, , sErrorText
    End If
End Sub

Public Function GetFreeChannel() As Long
    
    Dim bFreeChannelFound As Boolean
    Dim lChannelFreeStatusArrayIndex As Long
    Dim lFreeChannel As Long
    
    bFreeChannelFound = False
    
    For lChannelFreeStatusArrayIndex = LBound(mbaChannelFreeStatus) To UBound(mbaChannelFreeStatus)
        If mbaChannelFreeStatus(lChannelFreeStatusArrayIndex) = True Then
            lFreeChannel = lChannelFreeStatusArrayIndex
            bFreeChannelFound = True
            mbaChannelFreeStatus(lChannelFreeStatusArrayIndex) = False
            Exit For
        End If
    Next lChannelFreeStatusArrayIndex
    
    If Not bFreeChannelFound Then
        err.Raise 1000, , "Out of free channels"
    Else
        GetFreeChannel = lFreeChannel
    End If
    
End Function

Public Sub SetChannelStatus(ByVal vlChannelIndex As Long, ByVal vboolIsChannelFree As Boolean)
    mbaChannelFreeStatus(vlChannelIndex) = vboolIsChannelFree
End Sub

