VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MIDIEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const MMSYSERR_NOERROR = 0  '  no error

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Declare Function midiOutClose Lib "winmm.dll" (ByVal hMidiOut As Long) As Long
Private Declare Function midiOutOpen Lib "winmm.dll" (lphMidiOut As Long, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Private Declare Function midiOutShortMsg Lib "winmm.dll" (ByVal hMidiOut As Long, ByVal dwMsg As Long) As Long
Private Declare Function midiOutReset Lib "winmm.dll" (ByVal hMidiOut As Long) As Long
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function midiOutSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
Private Declare Function midiOutGetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, lpdwVolume As Long) As Long

Private mlMIDIOutHandle As Long
Private mbExternalHandle As Boolean
Private mbIsPlayingNote As Boolean

Private m_lOctave As Long
Private m_lInstrument As Long
Private m_lChannel As Long
Private m_lPan As Long
Private m_lVolume As Long
Private m_lLastNote As Long
Private m_lNote As Long
Private m_lLastOctave As Long

Public Event NoteStarted(ByVal vlNoteNumber As Long, ByVal vlOctave As Long, ByVal vlVolume As Long)
Public Event NoteStoped(ByVal vlNoteNumber As Long, ByVal vlOctave As Long)
Public Event PropertyChanged(ByVal venmProperty As MIDIEngineProperties, ByVal rlOldValue As Long, ByRef rlNewValue As Long, ByRef rbIgnoreNewValue As Boolean)

Public Sub OpenOutputPort(Optional ByVal vboolAllocateChannel As Boolean = False)
    
    mlMIDIOutHandle = GetMIDIOutHandle
    mbExternalHandle = False
    
    If vboolAllocateChannel And (m_lChannel = -1) Then
        m_lChannel = GetFreeChannel
    End If
    
End Sub

Private Sub SendMidiOut(ByVal vlEvent As Long, ByVal vlNote As Long, ByVal vlVelocity As Long)
      
    Dim lLowWord As Long
    Dim lHighWord As Long
    Dim lMidiAPIReturn As Long
    Dim lMidiMessage As Long
   
    lLowWord = (vlNote * 256) + vlEvent
    lHighWord = ((vlVelocity Mod 256) * 256) * 256

    lMidiMessage = lLowWord + lHighWord
   
    lMidiAPIReturn = midiOutShortMsg(mlMIDIOutHandle, lMidiMessage)
    Call CheckMidiApiReturn(lMidiAPIReturn)

End Sub

Public Sub PlayNote(Optional ByVal vlNote As Variant, Optional ByVal vlOctave As Variant = Empty, Optional ByVal vlVolume As Variant = Empty)
    
    Dim lMidiEvent As Long
    Dim lMidiNote As Long
    Dim lMidiOctave As Long
    Dim lMidiVolume As Long
    Dim lActualMidiNote As Long
    
    If IsEmpty(vlNote) Then
        lMidiNote = m_lNote
    Else
        lMidiNote = vlNote
    End If
    
    If IsEmpty(vlOctave) Then
        lMidiOctave = m_lOctave
    Else
        lMidiOctave = vlOctave
    End If
    
    If IsEmpty(vlVolume) Then
        lMidiVolume = m_lVolume
    Else
        lMidiVolume = vlVolume
    End If
    
    lActualMidiNote = lMidiNote + FirstMidiNoteNumberForOctave(lMidiOctave)
    
    If m_lChannel = -1 Then
        m_lChannel = GetFreeChannel
    End If
    
    lMidiEvent = 144 + m_lChannel
    Call SendMidiOut(lMidiEvent, lActualMidiNote, lMidiVolume)
    
    mbIsPlayingNote = True
    RaiseEvent NoteStarted(lMidiNote, lMidiOctave, lMidiVolume)
    m_lLastNote = lMidiNote
    m_lLastOctave = lMidiOctave
End Sub

Public Sub StopNote(Optional ByVal vlNote As Variant, Optional ByVal vlOctave As Variant)
    
    Dim lMidiEvent As Long
    Dim lMidiNote As Long
    Dim lMidiOctave As Long
    Dim lActualMidiNote As Long
    
    If IsMissing(vlNote) Or IsEmpty(vlNote) Then
        lMidiNote = m_lLastNote
    Else
        lMidiNote = vlNote
    End If
    
    If IsMissing(vlOctave) Or IsEmpty(vlOctave) Then
        lMidiOctave = m_lLastOctave
    Else
        lMidiOctave = vlOctave
    End If
    
    lMidiEvent = 144 + m_lChannel
    lActualMidiNote = lMidiNote + FirstMidiNoteNumberForOctave(lMidiOctave)
    Call SendMidiOut(lMidiEvent, lActualMidiNote, 0)
    mbIsPlayingNote = False
    RaiseEvent NoteStoped(lMidiNote, lMidiOctave)
End Sub

Private Function FirstMidiNoteNumberForOctave(ByVal vlOctave As Long) As Long
    FirstMidiNoteNumberForOctave = vlOctave * 12
End Function

Public Property Get Octave() As Long
    Octave = m_lOctave
End Property
Public Property Get Instrument() As Long
    Instrument = m_lInstrument
End Property
Public Property Get Channel() As Long
    If m_lChannel = -1 Then m_lChannel = GetFreeChannel
    Channel = m_lChannel
End Property
Public Property Get Pan() As Long
    Pan = m_lPan
End Property
Public Property Get Volume() As Long
    Volume = m_lVolume
End Property
Public Property Get LastNote() As Long
    LastNote = m_lLastNote
End Property
Public Property Get LastOctave() As Long
    LastOctave = m_lLastOctave
End Property
Public Property Get Handle() As Long
    Handle = mlMIDIOutHandle
End Property
Public Property Get Note() As Long
    Note = m_lNote
End Property

Public Property Let Note(ByVal vlNote As Long)
    Dim bIsPlayingNote As Boolean
    
    bIsPlayingNote = mbIsPlayingNote
    
    If bIsPlayingNote Then
        Call StopNote
    End If
    
    m_lNote = vlNote
    
    If bIsPlayingNote Then
        Call PlayNote
    End If
    
End Property
Public Property Let Octave(ByVal vlOctave As Long)
    If vlOctave <> m_lOctave Then
        Dim bAllowChange As Boolean
        bAllowChange = True
        RaiseEvent PropertyChanged(mepOctave, m_lOctave, vlOctave, bAllowChange)
        If bAllowChange = True Then
            m_lOctave = vlOctave
        End If
    End If
End Property
Public Property Let Handle(ByVal vlHandle As Long)
    mlMIDIOutHandle = vlHandle
    mbExternalHandle = True
End Property

Public Property Let Instrument(ByVal vlInstrument As Long)
    
   Dim lMidiEvent As Long
   Dim lMidiNote As Long
   Dim lMidiVolume As Long

    If m_lInstrument <> vlInstrument Or m_lChannel = -1 Then
        Dim bAllowChange As Boolean
        bAllowChange = True
        RaiseEvent PropertyChanged(mepInstrument, m_lInstrument, vlInstrument, bAllowChange)
        If bAllowChange = True Then
            m_lInstrument = vlInstrument
            
            If m_lChannel = -1 Then
                m_lChannel = GetFreeChannel
            End If
            
            lMidiEvent = &HC0 + m_lChannel
            lMidiNote = m_lInstrument
            lMidiVolume = 0
            
            Call SendMidiOut(lMidiEvent, lMidiNote, lMidiVolume)
        End If
    End If
    
End Property
Public Property Let Channel(ByVal vlChannel As Long)
    
    Dim bIsPlayingNote As Boolean
    
    bIsPlayingNote = mbIsPlayingNote
    
    If m_lChannel <> -1 Then
        
        If bIsPlayingNote Then
            Call StopNote
        End If
        
        Call SetChannelStatus(m_lChannel, True)
        
    End If
    
    m_lChannel = vlChannel
    
    Call SetChannelStatus(m_lChannel, False)
    
    If bIsPlayingNote Then
        Call PlayNote(m_lLastNote)
    End If
    
End Property
Public Property Let Pan(ByVal vlPan As Long)
  
   Dim lMidiEvent As Long
   Dim lMidiNote As Long
   Dim lMidiVolume As Long

    If m_lPan <> vlPan Or m_lChannel = -1 Then
    
        Dim bAllowChange As Boolean
        bAllowChange = True
        RaiseEvent PropertyChanged(mepPan, m_lPan, vlPan, bAllowChange)
        If bAllowChange = True Then
            m_lPan = vlPan
            
            If m_lChannel = -1 Then m_lChannel = GetFreeChannel
            
            lMidiEvent = 176 + m_lChannel
            lMidiNote = 10
            lMidiVolume = m_lVolume
            
            Call SendMidiOut(lMidiEvent, lMidiNote, lMidiVolume)
        End If
    End If
   
End Property
Public Property Let Volume(ByVal vlVolume As Long)
    If m_lVolume <> vlVolume Then
        Dim bAllowChange As Boolean
        bAllowChange = True
        RaiseEvent PropertyChanged(mepVolume, m_lVolume, vlVolume, bAllowChange)
        If bAllowChange = True Then
            m_lVolume = vlVolume
            'Call midiOutSetVolume(Handle, CLng((m_lVolume * &HFFFF) / 100))
        End If
    End If
End Property

Private Sub Class_Initialize()
    mbExternalHandle = False
    mbIsPlayingNote = False
    m_lChannel = -1 'Channel not allocated
    Call SetToDefaults
End Sub

Private Sub Class_Terminate()
    Call CloseOutputPort
End Sub

Public Sub CloseOutputPort()
    If (Not mbExternalHandle) And (mlMIDIOutHandle <> 0) Then
        Call ReleaseMIDIOutHandle
    End If
    
    If (m_lChannel <> -1) Then
        Call SetChannelStatus(m_lChannel, True)
    End If
End Sub

Public Function IsMidiOutPortOpen() As Boolean
    If mlMIDIOutHandle <> 0 Then
        IsMidiOutPortOpen = True
    Else
        IsMidiOutPortOpen = False
    End If
End Function

Public Function IsPlayingNote() As Boolean
    IsPlayingNote = mbIsPlayingNote
End Function

Public Sub ReleaseChannel()
    If m_lChannel <> -1 Then
        If mbIsPlayingNote Then
            Call StopNote
        End If
        Call SetChannelStatus(m_lChannel, True)
        m_lChannel = -1
    End If
End Sub

Public Sub SetToDefaults()
    m_lInstrument = lMIDI_ENG_DEFAULT_INSTRUMENT
    m_lOctave = lMIDI_ENG_DEFAULT_OCTAVE
    m_lPan = lMIDI_ENG_DEFAULT_PAN
    m_lVolume = lMIDI_ENG_DEFAULT_VOLUME
End Sub

Public Function NoteIndexToName(ByVal vlIndex As Long) As String
    Dim lModIndex As Long
    Dim sNoteName As String
    lModIndex = vlIndex Mod 12
    Select Case lModIndex
        Case 0:
            sNoteName = "C"
        Case 1:
            sNoteName = "C#"
        Case 2:
            sNoteName = "D"
        Case 3:
            sNoteName = "D#"
        Case 4:
            sNoteName = "E"
        Case 5:
            sNoteName = "F"
        Case 6:
            sNoteName = "F#"
        Case 7:
            sNoteName = "G"
        Case 8:
            sNoteName = "G#"
        Case 9:
            sNoteName = "A"
        Case 10:
            sNoteName = "A#"
        Case 11:
            sNoteName = "B"
    End Select
    
    NoteIndexToName = sNoteName
End Function

Public Sub StopAllNotes()
    Dim lMidiAPIReturn As Long
   
    lMidiAPIReturn = midiOutReset(mlMIDIOutHandle)
    'If no channel allocated then ignore the errors
    If m_lChannel <> -1 Then
        Call CheckMidiApiReturn(lMidiAPIReturn)
    End If
End Sub

Public Sub RaisePropertyChangeEvents()
    Dim bAllowChange As Boolean
    bAllowChange = True
    RaiseEvent PropertyChanged(mepOctave, m_lOctave, m_lOctave, bAllowChange)
    bAllowChange = True
    RaiseEvent PropertyChanged(mepInstrument, m_lInstrument, m_lInstrument, bAllowChange)
    bAllowChange = True
    RaiseEvent PropertyChanged(mepPan, m_lPan, m_lPan, bAllowChange)
    bAllowChange = True
    RaiseEvent PropertyChanged(mepVolume, m_lVolume, m_lVolume, bAllowChange)
End Sub
