VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_sName As String
Private m_sText As String
Private moclInstructions As Instructions
Private mboolCompiled As Boolean
Private moMScriptCompiler As MScriptCompiler
Private moclParentMScriptsCol As MScripts
Private WithEvents moInstructionProcesor As InstructionProcessor
Attribute moInstructionProcesor.VB_VarHelpID = -1
Private moInstructionClock As ClockProvider

Public Property Get InstructionProcessor() As InstructionProcessor
    Set InstructionProcessor = moInstructionProcesor
End Property

Public Property Get InstructionClock() As ClockProvider
    Set InstructionClock = moInstructionClock
End Property

Public Property Set InstructionClock(voInstructionClock As ClockProvider)
    Set moInstructionClock = voInstructionClock
End Property

Public Property Get MScriptCompiler() As MScriptCompiler
    Set MScriptCompiler = moMScriptCompiler
End Property

Public Property Set MScriptCompiler(voMScriptCompiler As MScriptCompiler)
    Set moMScriptCompiler = voMScriptCompiler
End Property

Public Property Get ParentMScriptsCol() As MScripts
    Set ParentMScriptsCol = moclParentMScriptsCol
End Property

Public Property Set ParentMScriptsCol(voParentMScriptsCol As MScripts)
    Set moclParentMScriptsCol = voParentMScriptsCol
End Property

Public Property Get Text() As String
Attribute Text.VB_UserMemId = 0
    Text = m_sText
End Property

Public Property Let Text(ByVal vsText As String)
    
    If StrComp(m_sText, vsText, vbBinaryCompare) <> 0 Then
    
        Dim bCanceled As Boolean
        bCanceled = False
        
        Call moclParentMScriptsCol.GenerateModifyMScriptEvent(m_sName, m_sText, vsText, bCanceled)
        
        If Not bCanceled Then
    
            m_sText = vsText
            mboolCompiled = False
        
        End If
    
    End If
        
End Property

Public Function IsCompiled() As String
    IsCompiled = mboolCompiled
End Function

Private Sub Class_Initialize()
    mboolCompiled = False
End Sub

Public Sub Compile(Optional ByVal vlCompileStartPoint As Long = 1, Optional ByVal vlCompileEndPoint As Long = -1)
    moclParentMScriptsCol.LastCompiledMScript = m_sName
    Set moclInstructions = Nothing
    Set moclInstructions = moMScriptCompiler.Compile(m_sText, vlCompileStartPoint, vlCompileEndPoint)
    If vlCompileStartPoint = 1 And vlCompileEndPoint = -1 Then
        mboolCompiled = True
    Else
        mboolCompiled = False
    End If
End Sub

Public Sub Play(Optional ByVal vlCompileStartPoint As Long = 1, Optional ByVal vlCompileEndPoint As Long = -1, Optional ByVal vsFromLable As String = "")
    If (Not mboolCompiled) Or (vlCompileStartPoint <> 1) Or (vlCompileEndPoint <> -1) Then
        Call Compile(vlCompileStartPoint, vlCompileEndPoint)
    End If
    If moInstructionProcesor Is Nothing Then
        Set moInstructionProcesor = New InstructionProcessor
        Set moInstructionProcesor.NoteClock = moInstructionClock
        Call moInstructionProcesor.moMIDIEngine.OpenOutputPort(True)
    End If
    If Not moInstructionProcesor.IsPaused Then
        Set moInstructionProcesor.Instructions = moclInstructions
    End If
    Call moInstructionProcesor.Execute(vsFromLable)
End Sub

Public Sub PausePlay()
    If Not moInstructionProcesor Is Nothing Then
        moInstructionProcesor.PauseExecution
    End If
End Sub

Public Sub StopPlay()
    If Not moInstructionProcesor Is Nothing Then
        moInstructionProcesor.StopExecution
    End If
End Sub

Private Sub Class_Terminate()
    Set moclInstructions = Nothing
    Set moclParentMScriptsCol = Nothing
    Set moInstructionProcesor = Nothing
    Set moMScriptCompiler = Nothing
End Sub

Private Sub moInstructionProcesor_Error(ByVal Number As Long, ByVal Source As String, ByVal Description As String, ByVal HelpFile As String, ByVal HelpContext As Long)
    Call moclParentMScriptsCol.GenerateInstructionProcessorErrorEvent(m_sName, Number, Source, Description, HelpFile, HelpContext)
End Sub

Private Sub moInstructionProcesor_ExecuteSub(ByVal venmAction As ExecuteSubActions, ByVal vsSubName As String, ByVal vsLableName As String)
    
    On Error GoTo ERR_moInstructionProcesor_ExecuteSub
    
    Select Case venmAction
        Case esaPlay
            Call moclParentMScriptsCol(vsSubName).Play(, , vsLableName)
        Case esaStop
            Call moclParentMScriptsCol(vsSubName).StopPlay
        Case esaPause
            Call moclParentMScriptsCol(vsSubName).PausePlay
    End Select
    
Exit Sub
ERR_moInstructionProcesor_ExecuteSub:
    Call moclParentMScriptsCol.GenerateInstructionProcessorErrorEvent(m_sName, err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext)
End Sub

Public Property Get Name() As String
    Name = m_sName
End Property

Friend Property Let Name(ByVal vsName As String)
    m_sName = vsName
End Property

Private Sub moInstructionProcesor_ExecutionStartStop(ByVal vboolStartStopFlag As Boolean)
    Call moclParentMScriptsCol.GenerateInstructionProcessorExecutionStartStopEvent(m_sName, vboolStartStopFlag)
End Sub

Private Sub moInstructionProcesor_PrintRequest(ByVal vsStringToPrint As String)
    Call moclParentMScriptsCol.GeneratePrintRequestEvent(m_sName, vsStringToPrint)
End Sub

Private Sub moInstructionProcesor_StateChanged(ByVal venmOldState As InstructionProcessorStatus, ByVal venmNewState As InstructionProcessorStatus)
    Call moclParentMScriptsCol.GenerateInstructionProcessorStateChangedEvent(m_sName, venmOldState, venmNewState)
End Sub
