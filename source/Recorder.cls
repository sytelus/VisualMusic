VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MScriptRecorder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents MIDIEngine As MIDIEngine
Attribute MIDIEngine.VB_VarHelpID = -1

Public WithEvents ClockProvider As ClockProvider
Attribute ClockProvider.VB_VarHelpID = -1

Private WithEvents SavedMIDIEngineWhenPaused As MIDIEngine
Private SavedClockProviderWhenPaused As ClockProvider

Private mbRecordMode As Boolean
Private mlNoteInterval As Long
Private mlSilenceInterval As Long
Private mbIsFirstNote As Boolean
Private mbNoteStarted As Boolean
Private m_sMScriptText As String
Private m_bEnableMultiKeyRecording As Boolean
Private mbIsPaused As Boolean
Private mbRecordPropertyChangeWhilePaused As Boolean

Public Event MScriptTextChanged(ByVal vboolIsAppended As Boolean, ByVal vsAppendedText As String)
Public Event RecordingStatusChanged(ByVal vboolIsRecording As Boolean)

Public Property Get EnableMultiKeyRecording() As Boolean
    EnableMultiKeyRecording = m_bEnableMultiKeyRecording
End Property

Public Property Let EnableMultiKeyRecording(ByVal vlEnableMultiKeyRecording As Boolean)
    m_bEnableMultiKeyRecording = vlEnableMultiKeyRecording
    If mbRecordMode Then
        mbNoteStarted = False
        mlNoteInterval = 0
        mlSilenceInterval = 0
    End If
End Property

Public Property Get MScriptText() As String
    MScriptText = m_sMScriptText
End Property

Private Function AppendToMScriptText(ByVal vsStringToAppend As String)
    m_sMScriptText = m_sMScriptText & vsStringToAppend
    RaiseEvent MScriptTextChanged(True, vsStringToAppend)
End Function

Public Function StartRecording(Optional ByVal vboolRestartIfRecording As Boolean = False, Optional vboolIncludeAllPropertiesInStartString As Boolean = False) As String
    If mbRecordMode = True Then
        If mbIsPaused Then
            Call UnPause
        Else
            If vboolRestartIfRecording = False Then
                err.Raise 1000, "", "Recording is already started"
            End If
        End If
    Else
        mbRecordMode = True
        mbIsFirstNote = True
        mbNoteStarted = False
        mbIsPaused = False
        m_sMScriptText = vbNullString
        With Me.MIDIEngine
            If .Instrument <> lMIDI_ENG_DEFAULT_INSTRUMENT _
                Or vboolIncludeAllPropertiesInStartString = True Then
                m_sMScriptText = m_sMScriptText & " I:" & .Instrument
            End If
            If .Volume <> lMIDI_ENG_DEFAULT_VOLUME _
                Or vboolIncludeAllPropertiesInStartString = True Then
                m_sMScriptText = m_sMScriptText & " V:" & .Volume
            End If
            If .Octave <> lMIDI_ENG_DEFAULT_OCTAVE _
                Or vboolIncludeAllPropertiesInStartString = True Then
                m_sMScriptText = m_sMScriptText & " O:" & .Octave
            End If
            If .Pan <> lMIDI_ENG_DEFAULT_PAN _
                Or vboolIncludeAllPropertiesInStartString = True Then
                m_sMScriptText = m_sMScriptText & " P:" & .Pan
            End If
        End With
    End If
    StartRecording = m_sMScriptText
    RaiseEvent RecordingStatusChanged(True)
End Function

Private Sub UnPause()
    If mbIsPaused Then
        mbIsPaused = False
        Set MIDIEngine = SavedMIDIEngineWhenPaused
        Set ClockProvider = SavedClockProviderWhenPaused
        Set SavedClockProviderWhenPaused = Nothing
        Set SavedMIDIEngineWhenPaused = Nothing
    End If
End Sub

Public Sub StopRecording(Optional ByVal vboolIncludeLastSilenceInterval As Boolean = False)
    mbRecordMode = False
    Call UnPause
    mbNoteStarted = False
    If Me.MIDIEngine.IsPlayingNote Then
        Call AppendToMScriptText(" " & Me.MIDIEngine.NoteIndexToName(Me.MIDIEngine.LastNote) & ":" & mlNoteInterval)
    End If
    If vboolIncludeLastSilenceInterval And (Not Me.MIDIEngine.IsPlayingNote) Then
        Call AppendToMScriptText(" " & "S:" & mlSilenceInterval)
    End If
    RaiseEvent RecordingStatusChanged(False)
End Sub

Public Sub PauseRecording(Optional ByVal vboolRecordPropertyChangeWhilePaused As Boolean = False, Optional ByVal vboolRestartRecordingIfPaused As Boolean = False)
    mbRecordPropertyChangeWhilePaused = vboolRecordPropertyChangeWhilePaused
    If mbIsPaused Then
        If vboolRestartRecordingIfPaused Then
            Call UnPause
        Else
            err.Raise 1000, , "Recording is already paused"
        End If
    Else
        mbIsPaused = True
        Set SavedClockProviderWhenPaused = ClockProvider
        Set SavedMIDIEngineWhenPaused = MIDIEngine
        Set ClockProvider = Nothing
        Set MIDIEngine = Nothing
    End If
End Sub

Public Function IsPaused() As Boolean
    IsPaused = mbIsPaused
End Function
Public Function IsRecording() As Boolean
    IsRecording = mbRecordMode
End Function

Private Sub Class_Initialize()
    mbRecordMode = False
    mbIsPaused = False
End Sub

Private Sub Class_Terminate()
    If mbRecordMode Then
        Call StopRecording
    End If
    Set MIDIEngine = Nothing
    Set ClockProvider = Nothing
    Set SavedClockProviderWhenPaused = Nothing
    Set SavedMIDIEngineWhenPaused = Nothing
End Sub

Private Sub ClockProvider_ClockPulse()
    If mbRecordMode = True Then
        mlNoteInterval = mlNoteInterval + 1
        mlSilenceInterval = mlSilenceInterval + 1
    End If
End Sub

Private Sub MIDIEngine_NoteStarted(ByVal vlNoteNumber As Long, ByVal vlOctave As Long, ByVal vlVolume As Long)
    If mbRecordMode Then
        If Not m_bEnableMultiKeyRecording Then
            If mbNoteStarted Then
                Call MIDIEngine_NoteStoped(Me.MIDIEngine.LastNote, Me.MIDIEngine.LastOctave)
            End If
            mlNoteInterval = 0
            mbNoteStarted = True
            If Not mbIsFirstNote Then
                Call AppendToMScriptText(" " & "s:" & mlSilenceInterval)
            End If
        Else
            If Not mbIsFirstNote Then
                Call AppendToMScriptText(" " & "s:" & mlSilenceInterval)
            End If
            Call AppendToMScriptText(" " & "NON:" & vlNoteNumber)
            mlSilenceInterval = 0
        End If
    End If
End Sub

Private Sub MIDIEngine_NoteStoped(ByVal vlNoteNumber As Long, ByVal vlOctave As Long)
    If mbRecordMode Then
        If Not m_bEnableMultiKeyRecording Then
            mbNoteStarted = False
            mbIsFirstNote = False
            Call AppendToMScriptText(" " & GetMScriptForNote(vlNoteNumber, vlOctave, mlNoteInterval))
            mlSilenceInterval = 0
        Else
            mbIsFirstNote = False
            Call AppendToMScriptText(" " & "s:" & mlSilenceInterval)
            Call AppendToMScriptText(" " & "NOFF:" & vlNoteNumber)
            mlSilenceInterval = 0
        End If
    End If
End Sub

Private Sub MIDIEngine_PropertyChanged(ByVal venmProperty As MIDIEngineProperties, ByVal rlOldValue As Long, rlNewValue As Long, rbIgnoreNewValue As Boolean)
    If mbRecordMode Then
        Dim sInstructionCode As String
        
        Select Case venmProperty
            Case mepInstrument: sInstructionCode = "I"
            Case mepVolume: sInstructionCode = "V"
            Case mepPan: sInstructionCode = "P"
            Case mepOctave: sInstructionCode = "O"
        End Select
        
        Call AppendToMScriptText(" " & sInstructionCode & ":" & rlNewValue)
    End If
End Sub

Private Function GetMScriptForNote(ByVal vlNoteNumber As Long, ByVal vlOctave As Long, Optional ByVal vlNoteDuration As Variant, Optional ByVal vlSilenceDuration As Variant, _
    Optional ByVal vlVolume As Variant) As String
    
    Dim sMScrtiptTextForNote As String
    
    sMScrtiptTextForNote = Me.MIDIEngine.NoteIndexToName(vlNoteNumber) & ":"
    
    If Not IsMissing(vlNoteDuration) Then
        sMScrtiptTextForNote = sMScrtiptTextForNote & vlNoteDuration
    End If
    sMScrtiptTextForNote = sMScrtiptTextForNote & ","

    Dim bSilenceDurationRequired As Boolean
    Dim bOctaveRequired As Boolean
    Dim bVolumeRequired As Boolean

    bSilenceDurationRequired = Not IsMissing(vlSilenceDuration)
    
    Dim lNoteOctaveNumber As Long
    lNoteOctaveNumber = vlNoteNumber \ 12 + vlOctave
    If Me.MIDIEngine.Octave <> lNoteOctaveNumber Then
        bOctaveRequired = True
    Else
        bOctaveRequired = False
    End If
    
    bVolumeRequired = Not IsMissing(vlVolume)
    
    If bSilenceDurationRequired Or bOctaveRequired Or bVolumeRequired Then
    
        If bSilenceDurationRequired Then
            sMScrtiptTextForNote = sMScrtiptTextForNote & vlSilenceDuration
        End If
        sMScrtiptTextForNote = sMScrtiptTextForNote & ","
        
        If bOctaveRequired Or bVolumeRequired Then
        
            If bOctaveRequired = True Then
                If lNoteOctaveNumber > Me.MIDIEngine.Octave Then
                    sMScrtiptTextForNote = sMScrtiptTextForNote & "++" & (lNoteOctaveNumber - Me.MIDIEngine.Octave)
                Else
                    sMScrtiptTextForNote = sMScrtiptTextForNote & "--" & (Me.MIDIEngine.Octave - lNoteOctaveNumber)
                End If
            End If
        
            If bVolumeRequired Then
                sMScrtiptTextForNote = sMScrtiptTextForNote & ","
                sMScrtiptTextForNote = sMScrtiptTextForNote & vlVolume
            End If
            
        End If

    End If
    
    GetMScriptForNote = sMScrtiptTextForNote
End Function

Private Sub SavedMIDIEngineWhenPaused_PropertyChanged(ByVal venmProperty As MIDIEngineProperties, ByVal rlOldValue As Long, rlNewValue As Long, rbIgnoreNewValue As Boolean)
    If mbRecordPropertyChangeWhilePaused Then
        Call MIDIEngine_PropertyChanged(venmProperty, rlOldValue, rlNewValue, rbIgnoreNewValue)
    End If
End Sub
