VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InstructionProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_oclInstructions As Instructions
Private m_lInstructionCounter As Long

Private mlNotePassedInterval As Long
Private mlNoteTotalInterval As Long
Private mlSilencePassedInterval As Long
Private mlSilenceTotalInterval As Long
Private mbIsNoteManual As Boolean

Private mbIsExecuting As Boolean
Private mbIsPaused As Boolean
Private m_enmState As InstructionProcessorStatus
Private m_lDefaultNoteInterval As Long
Private m_lDefaultSilenceInterval As Long
Private m_lSilenceIntervalMultiplier As Long
Private m_lNoteIntervalMultiplier As Long

Private moclGlobalVars As Collection

Public moMIDIEngine As MIDIEngine

Public Event ExecuteSub(ByVal venmAction As ExecuteSubActions, ByVal vsSubName As String, ByVal vsLableName As String)
Public Event StateChanged(ByVal venmOldState As InstructionProcessorStatus, ByVal venmNewState As InstructionProcessorStatus)
Public Event Error(ByVal Number As Long, ByVal Source As String, ByVal Description As String, ByVal HelpFile As String, ByVal HelpContext As Long)
Public Event PrintRequest(ByVal vsStringToPrint As String)
Public Event ExecutionStartStop(ByVal vboolStartStopFlag As Boolean)

Private Const sMODULE_NAME As String = "InstructionProcessor"
Private Const mlSILENCE_LEN_FOR_CHENNEL_RELEASE As Long = 300
Private Const lDEFAULT_RANDOM_UPPER_LIMIT As Long = 100

Private WithEvents m_oNoteClock As ClockProvider
Attribute m_oNoteClock.VB_VarHelpID = -1

Public Property Get SilenceIntervalMultiplier() As Long
    SilenceIntervalMultiplier = m_lSilenceIntervalMultiplier
End Property

Public Property Let SilenceIntervalMultiplier(ByVal vlIntervalMultiplier As Long)
    m_lSilenceIntervalMultiplier = vlIntervalMultiplier
End Property

Public Property Get NoteIntervalMultiplier() As Long
    NoteIntervalMultiplier = m_lNoteIntervalMultiplier
End Property

Public Property Let NoteIntervalMultiplier(ByVal vlIntervalMultiplier As Long)
    m_lNoteIntervalMultiplier = vlIntervalMultiplier
End Property

Public Property Get Instructions() As Instructions
    Set Instructions = m_oclInstructions
End Property
Public Property Set Instructions(ByVal voclInstructions As Instructions)
    Set m_oclInstructions = voclInstructions
End Property
Public Property Get InstructionCounter() As Long
    InstructionCounter = m_lInstructionCounter
End Property
Public Property Let InstructionCounter(ByVal vlInstructionCounter As Long)
    m_lInstructionCounter = vlInstructionCounter
End Property

Public Sub ResetInstructionCounter()
    m_lInstructionCounter = m_oclInstructions.LOWER_BOUND
End Sub

Public Sub Execute(Optional ByVal vsFromLable As String = "")
    
    mbIsExecuting = True
    RaiseEvent ExecutionStartStop(mbIsExecuting)
    
    moMIDIEngine.SetToDefaults
    m_lNoteIntervalMultiplier = 1
    m_lSilenceIntervalMultiplier = 1
    If Not mbIsPaused Then
        If vsFromLable = vbNullString Then
            Call ResetInstructionCounter
            m_lInstructionCounter = m_lInstructionCounter - 1
            Call ExecuteNextInstruction
        Else
            Call JumpToLable(vsFromLable)
        End If
    Else
        Call ExecuteNextInstruction
        mbIsPaused = False
    End If
End Sub

Public Function ExecuteNextInstruction() As Boolean

    On Error GoTo ERR_ExecuteNextInstruction

    If mbIsExecuting Then
        'Is invalid instruction count?
        If m_lInstructionCounter >= (m_oclInstructions.Count - m_oclInstructions.LOWER_BOUND + 1) Then
            
            ExecuteNextInstruction = False
            State = ipsIdle
            
            mbIsExecuting = False
            RaiseEvent ExecutionStartStop(mbIsExecuting)
            
        Else
            
            ExecuteNextInstruction = True
            
            m_lInstructionCounter = m_lInstructionCounter + 1
            
            Call ExecuteCurrentInstruction
            
        End If
    End If
    
Exit Function
ERR_ExecuteNextInstruction:
    RaiseEvent Error(err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext)
End Function

Public Sub ExecuteCurrentInstruction(Optional ByVal vboolIncrementInstructionCounter As Long = True)
    
    Dim bOldExecutingStatus As Boolean
    
    bOldExecutingStatus = mbIsExecuting
    
    mbIsExecuting = True
    
        
    'Get instruction info
    Dim vaInstructionInfo As Variant
    vaInstructionInfo = m_oclInstructions(m_lInstructionCounter)
    
    'Get the instruction code
    Dim lInstructionCode As Long
    Dim vParamValue As Variant
    Dim sJumpLable As String
    Dim sVarName As String
    Dim vVarValue As Variant
    Dim sSubName As String
    Dim sLableName As String
    
    Dim vNote As Variant
    Dim vOctave As Variant
    Dim vVolume As Variant
    
    lInstructionCode = vaInstructionInfo(glARRAY_INDEX_INSTRUCTION_CODE)
    
    Select Case lInstructionCode
    
        Case glINSTRUCTION_PLAY_NOTE
        
            vOctave = Empty
            vVolume = Empty
                        
            mlNotePassedInterval = 0
            If GetInstructionParam(vaInstructionInfo, glPLAY_NOTE_PARAM_NOTE_INTERVAL, vParamValue) Then
                mlNoteTotalInterval = vParamValue
            Else
                mlNoteTotalInterval = m_lDefaultNoteInterval
            End If
            
            mlSilencePassedInterval = 0
            If GetInstructionParam(vaInstructionInfo, glPLAY_NOTE_PARAM_SILENCE_INTERVAL, vParamValue) Then
                mlSilenceTotalInterval = vParamValue
            Else
                mlSilenceTotalInterval = m_lDefaultSilenceInterval
            End If
            
            If GetInstructionParam(vaInstructionInfo, glPLAY_NOTE_PARAM_NOTE, vParamValue, True, moMIDIEngine.LastNote) Then vNote = vParamValue
            If GetInstructionParam(vaInstructionInfo, glPLAY_NOTE_PARAM_OCTAVE, vParamValue, True, moMIDIEngine.Octave) Then vOctave = vParamValue
            If GetInstructionParam(vaInstructionInfo, glPLAY_NOTE_PARAM_VOLUME, vParamValue, True, moMIDIEngine.Volume) Then vVolume = vParamValue
'            If GetInstructionParam(vaInstructionInfo, glPLAY_NOTE_PARAM_INSTRUMENT, vParamValue, True) Then moMIDIEngine.Instrument = vParamValue
'            If GetInstructionParam(vaInstructionInfo, glPLAY_NOTE_PARAM_PAN, vParamValue, True) Then moMIDIEngine.Pan = vParamValue
            
            Call moMIDIEngine.PlayNote(vNote, vOctave, vVolume)
            mbIsNoteManual = False
            State = ipsPlayingNote
                                        
        Case glINSTRUCTION_SET_INSTRUMENT
            If GetInstructionParam(vaInstructionInfo, glARRAY_INDEX_FIRST_PARAM, vParamValue, True) Then
                moMIDIEngine.Instrument = vParamValue
            Else
                err.Raise 1000, , "No instrument number supplied"
            End If
            Call ExecuteNextInstruction
        
        Case glINSTRUCTION_SET_OCTAVE
            If GetInstructionParam(vaInstructionInfo, glARRAY_INDEX_FIRST_PARAM, vParamValue, True) Then moMIDIEngine.Octave = vParamValue
            Call ExecuteNextInstruction
    
        Case glINSTRUCTION_SET_PAN
            If GetInstructionParam(vaInstructionInfo, glARRAY_INDEX_FIRST_PARAM, vParamValue, True) Then moMIDIEngine.Pan = vParamValue
            Call ExecuteNextInstruction
    
        Case glINSTRUCTION_SET_VOLUME
            If GetInstructionParam(vaInstructionInfo, glARRAY_INDEX_FIRST_PARAM, vParamValue, True) Then moMIDIEngine.Volume = vParamValue
            Call ExecuteNextInstruction
        
        Case glINSTRUCTION_PLAY_SILENCE
            State = ipsPlayingSilence
            mlSilencePassedInterval = 0
            mlSilenceTotalInterval = 0
            If GetInstructionParam(vaInstructionInfo, glARRAY_INDEX_FIRST_PARAM, vParamValue, True) Then mlSilenceTotalInterval = vParamValue
            If mlSilenceTotalInterval >= mlSILENCE_LEN_FOR_CHENNEL_RELEASE Then
                Call moMIDIEngine.ReleaseChannel
            End If
                
        Case glINSTRUCTION_JUMP
            sJumpLable = vbNullString
            If GetInstructionParam(vaInstructionInfo, glARRAY_INDEX_FIRST_PARAM, vParamValue) Then sJumpLable = vParamValue
            Call JumpToLable(sJumpLable)
            
        Case glINSTRUCTION_SET_VAR
            If GetInstructionParam(vaInstructionInfo, glSET_VAR_PARAM_VAR_NAME, vParamValue) Then sVarName = vParamValue
            If GetInstructionParam(vaInstructionInfo, glSET_VAR_PARAM_VAR_VALUE, vParamValue) Then
                vVarValue = vParamValue
            Else
                vVarValue = Empty
            End If
            
            Call SetInstructionProcessorGlobalVar(sVarName, vVarValue)
            Call ExecuteNextInstruction
        
        Case glINSTRUCTION_IF
            Dim sIfConditionString As String
            Dim sIfOnFalseJumpLable As String
            Dim bIfConditionValue As Boolean
            Dim vIfCompareToValue As Variant
            
            If GetInstructionParam(vaInstructionInfo, glIF_PARAM_VAR_NAME, vParamValue) Then sVarName = vParamValue
            If GetInstructionParam(vaInstructionInfo, glIF_PARAM_CONDITION, vParamValue) Then sIfConditionString = vParamValue
            If GetInstructionParam(vaInstructionInfo, glIF_PARAM_VALUE, vParamValue) Then vIfCompareToValue = vParamValue
            If GetInstructionParam(vaInstructionInfo, glIF_PARAM_ON_FALSE_JUMP, vParamValue) Then sIfOnFalseJumpLable = vParamValue
            If GetInstructionParam(vaInstructionInfo, glIF_PARAM_ON_TRUE_JUMP, vParamValue) Then sJumpLable = vParamValue
            
            vVarValue = GetInstructionProcessorGlobalVarValue(sVarName)
            
            If IsValidNumber(vIfCompareToValue) Then
                vIfCompareToValue = Val(vIfCompareToValue)
            End If
            
            Select Case sIfConditionString
                Case ">=", "=>"
                    bIfConditionValue = (vVarValue >= vIfCompareToValue)
                Case ">"
                    bIfConditionValue = (vVarValue > vIfCompareToValue)
                Case "<"
                    bIfConditionValue = (vVarValue < vIfCompareToValue)
                Case "<=", "=<"
                    bIfConditionValue = (vVarValue <= vIfCompareToValue)
                Case "="
                    bIfConditionValue = (vVarValue = vIfCompareToValue)
                Case Else
                    err.Raise 1000, , "Unrecognised condtion in IF"
                End Select
            If bIfConditionValue = True Then
                Call JumpToLable(sJumpLable)
            ElseIf sIfOnFalseJumpLable <> vbNullString Then
                Call JumpToLable(sIfOnFalseJumpLable)
            Else
                Call ExecuteNextInstruction
            End If
            
        Case glINSTRUCTION_ARITHMATIC
            Dim sArithmaticOperation As String
            Dim vOperand1 As Variant
            Dim vOperand2 As Variant
        
            If GetInstructionParam(vaInstructionInfo, glARITHMATIC_PARAM_VAR_NAME, vParamValue) Then sVarName = vParamValue
            If GetInstructionParam(vaInstructionInfo, glARITHMATIC_PARAM_OPERATION, vParamValue) Then sArithmaticOperation = vParamValue
            If GetInstructionParam(vaInstructionInfo, glARITHMATIC_PARAM_OPERAND1, vParamValue, True) Then vOperand1 = vParamValue
            If GetInstructionParam(vaInstructionInfo, glARITHMATIC_PARAM_OPERAND2, vParamValue, True) Then vOperand2 = vParamValue
            
            vVarValue = GetInstructionProcessorGlobalVarValue(sVarName)
            
            If Not IsValidNumber(vOperand1) Then
                If Not IsEmpty(vOperand1) Then
                    Call SetInstructionProcessorGlobalVar(vOperand1, vOperand1)
                End If
            End If
            
            If Not IsValidNumber(vOperand2) Then
                If Not IsEmpty(vOperand2) Then
                    Call SetInstructionProcessorGlobalVar(vOperand2, vOperand2)
                End If
            End If
            
            Select Case sArithmaticOperation
                Case "++"
                    Call SetInstructionProcessorGlobalVar(sVarName, Val(vVarValue) + 1)
                Case "--"
                    Call SetInstructionProcessorGlobalVar(sVarName, Val(vVarValue) - 1)
                Case "+="
                    Call SetInstructionProcessorGlobalVar(sVarName, Val(vVarValue) + Val(vOperand1))
                Case "-="
                    Call SetInstructionProcessorGlobalVar(sVarName, Val(vVarValue) - Val(vOperand1))
                Case "*="
                    Call SetInstructionProcessorGlobalVar(sVarName, Val(vVarValue) * Val(vOperand1))
                Case "/*"
                    Call SetInstructionProcessorGlobalVar(sVarName, Val(vVarValue) / Val(vOperand1))
                Case "^"
                    Call SetInstructionProcessorGlobalVar(sVarName, Val(vOperand1) ^ Val(vOperand2))
                Case "+"
                    Call SetInstructionProcessorGlobalVar(sVarName, Val(vOperand1) + Val(vOperand2))
                Case "-"
                    Call SetInstructionProcessorGlobalVar(sVarName, Val(vOperand1) - Val(vOperand2))
                Case "*"
                    Call SetInstructionProcessorGlobalVar(sVarName, Val(vOperand1) * Val(vOperand2))
                Case "/"
                    Call SetInstructionProcessorGlobalVar(sVarName, Val(vOperand1) / Val(vOperand2))
                Case "\"
                    Call SetInstructionProcessorGlobalVar(sVarName, Val(vOperand1) \ Val(vOperand2))
                Case Else
                    err.Raise 1000, , "Unrecognised arithmatic operation"
            End Select
            Call ExecuteNextInstruction
            
        Case glINSTRUCTION_SET_NOTE_INTERVAL
            If GetInstructionParam(vaInstructionInfo, glARRAY_INDEX_FIRST_PARAM, vParamValue, True) Then m_lDefaultNoteInterval = vParamValue
            Call ExecuteNextInstruction
        
        Case glINSTRUCTION_SET_SILENCE_INTERVAL
            If GetInstructionParam(vaInstructionInfo, glARRAY_INDEX_FIRST_PARAM, vParamValue, True) Then m_lDefaultSilenceInterval = vParamValue
            Call ExecuteNextInstruction
        Case glINSTRUCTION_PLAY_SUB
            sLableName = vbNullString
            
            If GetInstructionParam(vaInstructionInfo, glPLAY_SUB_PARAM_SUB_NAME, vParamValue, False) Then sSubName = vParamValue
            If GetInstructionParam(vaInstructionInfo, glPLAY_SUB_PARAM_LABLE_NAME, vParamValue, False) Then sLableName = vParamValue
            
            RaiseEvent ExecuteSub(esaPlay, sSubName, sLableName)
            Call ExecuteNextInstruction
        Case glINSTRUCTION_PAUSE_SUB
            If GetInstructionParam(vaInstructionInfo, glARRAY_INDEX_FIRST_PARAM, vParamValue, False) Then
                sSubName = vParamValue
                RaiseEvent ExecuteSub(esaPause, sSubName, vbNullString)
                Call ExecuteNextInstruction
            Else
                Call Me.PauseExecution
            End If
        Case glINSTRUCTION_STOP_SUB
            If GetInstructionParam(vaInstructionInfo, glARRAY_INDEX_FIRST_PARAM, vParamValue, False) Then
                sSubName = vParamValue
                RaiseEvent ExecuteSub(esaStop, sSubName, vbNullString)
                Call ExecuteNextInstruction
            Else
                Call Me.StopExecution
            End If
        Case glINSTRUCTION_RELEASE_CHANNEL
            moMIDIEngine.ReleaseChannel
            Call ExecuteNextInstruction
        Case glINSTRUCTION_NOTE_START
            If GetInstructionParam(vaInstructionInfo, glNOTE_START_STOP_PARAM_NOTE_NUMBER, vParamValue, True) Then vNote = vParamValue
            If GetInstructionParam(vaInstructionInfo, glNOTE_START_STOP_PARAM_OCTAVE, vParamValue, True) Then vOctave = vParamValue
            Call moMIDIEngine.PlayNote(vNote, vOctave)
            State = ipsPlayingNote
            mbIsNoteManual = True
            Call ExecuteNextInstruction
        Case glINSTRUCTION_NOTE_STOP
            If GetInstructionParam(vaInstructionInfo, glNOTE_START_STOP_PARAM_NOTE_NUMBER, vParamValue, True) Then vNote = vParamValue
            If GetInstructionParam(vaInstructionInfo, glNOTE_START_STOP_PARAM_OCTAVE, vParamValue, True) Then vOctave = vParamValue
            Call moMIDIEngine.StopNote(vNote, vOctave)
            Call ExecuteNextInstruction
        Case glINSTRUCTION_NO_OPERATION
            Call ExecuteNextInstruction
        Case glINSTRUCTION_TEMPO
            If GetInstructionParam(vaInstructionInfo, glARRAY_INDEX_FIRST_PARAM, vParamValue, True) Then
                m_lNoteIntervalMultiplier = vParamValue
                m_lSilenceIntervalMultiplier = vParamValue
            End If
            Call ExecuteNextInstruction
        Case glINSTRUCTION_RANDOM
            Dim lRandomUpperLimit As Long
            If GetInstructionParam(vaInstructionInfo, glRANDOM_PARAM_VAR_NAME, vParamValue) Then sVarName = vParamValue
            If GetInstructionParam(vaInstructionInfo, glSET_VAR_PARAM_VAR_VALUE, vParamValue) Then
                lRandomUpperLimit = vParamValue
            Else
                lRandomUpperLimit = lDEFAULT_RANDOM_UPPER_LIMIT
            End If
            Randomize
            Call SetInstructionProcessorGlobalVar(sVarName, CLng(Rnd * lRandomUpperLimit))
            Call ExecuteNextInstruction
        Case glINSTRUCTION_PRINT
            Dim sStringToPrint As String
            Dim sTagWord As String
            Dim sVarOrParamName As String
            
            If GetInstructionParam(vaInstructionInfo, glPRINT_PARAM_VAR_OR_PARAM_NAME, vParamValue) Then sVarOrParamName = vParamValue
            If GetInstructionParam(vaInstructionInfo, glPRINT_PARAM_TAG_WORD, vParamValue) Then sTagWord = vParamValue
                    
            Select Case sVarOrParamName
                Case "NI"
                    sStringToPrint = m_lNoteIntervalMultiplier
                Case "SI"
                    sStringToPrint = m_lSilenceIntervalMultiplier
                Case "OCTAVE"
                    sStringToPrint = moMIDIEngine.Octave
                Case "VOLUME"
                    sStringToPrint = moMIDIEngine.Volume
                Case "PAN"
                    sStringToPrint = moMIDIEngine.Pan
                Case Else
                    If Trim$(sVarOrParamName) <> vbNullString Then
                        sStringToPrint = GetInstructionProcessorGlobalVarValue(sVarOrParamName)
                    Else
                        sStringToPrint = vbNullString
                    End If
            End Select
            
            If sTagWord <> vbNullString Then
                sStringToPrint = sTagWord & IIf(sStringToPrint <> vbNullString, ", ", vbNullString) & sStringToPrint
            End If
                        
            RaiseEvent PrintRequest(sStringToPrint)
            Call ExecuteNextInstruction
        Case Else
            
            err.Raise 1000, , "Unknown instruction code"
            
    End Select
    
    mbIsExecuting = bOldExecutingStatus
    
End Sub

Private Sub Class_Initialize()
    m_lNoteIntervalMultiplier = 1
    m_lSilenceIntervalMultiplier = 1
    Set moMIDIEngine = New MIDIEngine
    State = ipsIdle
    Set moclGlobalVars = GetInstructionProcessorGlobalVarCol
    mbIsPaused = False
    m_lDefaultNoteInterval = 3
    m_lDefaultSilenceInterval = 0
    mbIsExecuting = False
    RaiseEvent ExecutionStartStop(mbIsExecuting)
End Sub

Private Function GetInstructionParam(ByVal vvaInstructionInfo As Variant, ByVal vlParamIndex As Long, ByRef rvParamValue As Variant, Optional ByVal vboolIsNumeric As Boolean = False, Optional vlCurrenNumericValue As Variant) As Boolean
    
    On Error Resume Next
    
    rvParamValue = Empty
    
    rvParamValue = vvaInstructionInfo(vlParamIndex)
    
    If (err.Number = 0) And (Not IsEmpty(rvParamValue)) Then
        rvParamValue = rvParamValue & vbNullString 'Null value fix
        If rvParamValue <> vbNullString Then
            If vboolIsNumeric Then
                If Not IsValidNumber(rvParamValue) Then
                    If Not IsMissing(vlCurrenNumericValue) Then
                        Dim sOperationCode As String
                        Dim sOperand As String
                        sOperationCode = Mid$(rvParamValue, 1, 2)
                        sOperand = Mid$(rvParamValue, 3)
                        Select Case CStr(sOperationCode)
                            Case "++":
                                rvParamValue = vlCurrenNumericValue + Val(sOperand)
                            Case "--":
                                rvParamValue = vlCurrenNumericValue - Val(sOperand)
                            Case "**":
                                rvParamValue = vlCurrenNumericValue * Val(sOperand)
                            Case "//":
                                rvParamValue = vlCurrenNumericValue / Val(sOperand)
                            Case Else
                                rvParamValue = Val(GetInstructionProcessorGlobalVarValue(rvParamValue))
                        End Select
                    Else
                        rvParamValue = Val(GetInstructionProcessorGlobalVarValue(rvParamValue))
                    End If
                End If
            End If
            If err.Number = 0 Then
                GetInstructionParam = True
            Else
                ReRaiseError
                GetInstructionParam = False
            End If
        Else
            GetInstructionParam = False
        End If
    Else
        GetInstructionParam = False
    End If

End Function

Private Sub Class_Terminate()
    If mbIsExecuting = True Then
        Call StopExecution
    End If
    Set moMIDIEngine = Nothing
    Set moclGlobalVars = Nothing
    Call ReleaseInstructionProcessorGlobalVarCol
End Sub

Public Function IsExecuting() As Boolean
    IsExecuting = mbIsExecuting
End Function

Public Property Get State() As InstructionProcessorStatus
    State = m_enmState
End Property

Private Property Let State(ByVal venmInstrunctionProcessorStatus As InstructionProcessorStatus)
    If m_enmState <> venmInstrunctionProcessorStatus Then
        RaiseEvent StateChanged(m_enmState, venmInstrunctionProcessorStatus)
        m_enmState = venmInstrunctionProcessorStatus
        If m_enmState = ipsIdle Then
            Call moMIDIEngine.ReleaseChannel
        End If
    End If
End Property

Public Property Get NoteClock() As ClockProvider
    Set NoteClock = m_oNoteClock
End Property

Public Property Set NoteClock(ByVal voNoteClock As ClockProvider)
    Set m_oNoteClock = voNoteClock
End Property

Private Sub m_oNoteClock_ClockPulse()
    If Not mbIsPaused Then
        If (m_enmState = ipsPlayingNote) = True Then
            If Not mbIsNoteManual Then
                mlNotePassedInterval = mlNotePassedInterval + 1
                If mlNotePassedInterval >= (mlNoteTotalInterval * m_lNoteIntervalMultiplier) Then
                    Call moMIDIEngine.StopNote
                    If mlSilenceTotalInterval = 0 Then
                        State = ipsExecutingCommand
                        Call ExecuteNextInstruction
                    Else
                        State = ipsPlayingSilence
                        mlSilencePassedInterval = 0
                    End If
                End If
            End If
        ElseIf (m_enmState = ipsPlayingSilence) Then
            mlSilencePassedInterval = mlSilencePassedInterval + 1
            If mlSilencePassedInterval >= (mlSilenceTotalInterval * m_lSilenceIntervalMultiplier) Then
                State = ipsExecutingCommand
                Call ExecuteNextInstruction
            End If
        End If
    End If
End Sub

Private Sub JumpToLable(ByVal vsLable As String)
    m_lInstructionCounter = GetInstructionNumberForLable(vsLable)
    m_lInstructionCounter = m_lInstructionCounter - 1
    Call ExecuteNextInstruction
End Sub

Public Function IsPaused() As Boolean
     IsPaused = mbIsPaused
End Function

Public Sub PauseExecution(Optional ByVal vboolStopNote As Boolean = True)
    If vboolStopNote And m_enmState = ipsPlayingNote Then
        moMIDIEngine.StopAllNotes
    End If
    mbIsPaused = True
End Sub

Public Sub StopExecution()
    State = ipsIdle
    mbIsExecuting = False
    moMIDIEngine.StopAllNotes
    RaiseEvent ExecutionStartStop(mbIsExecuting)
End Sub

Public Property Get DefaultNoteInterval() As Long
    DefaultNoteInterval = m_lDefaultNoteInterval
End Property

Public Property Get DefaultSilenceInterval() As Long
    DefaultSilenceInterval = m_lDefaultSilenceInterval
End Property

Public Property Let DefaultNoteInterval(ByVal vlInterval As Long)
    m_lDefaultNoteInterval = vlInterval
End Property

Public Property Let DefaultSilenceInterval(ByVal vlInterval As Long)
    m_lDefaultSilenceInterval = vlInterval
End Property

Private Function GetInstructionNumberForLable(ByVal vsLable As String) As Long
    On Error GoTo ERR_GetInstructionNumberForLable
    GetInstructionNumberForLable = m_oclInstructions.Lables(vsLable)
Exit Function
ERR_GetInstructionNumberForLable:
    If err.Number = 5 Then
        err.Raise lERR_RUN_INVALID_LABLE_NAME, , "Lable " & AlternateStrIfNull(vsLable, "<Lable name not available>") & " do not exist"
    Else
        ReRaiseError
    End If
End Function
