Attribute VB_Name = "InstructionProcessorGlobals"
Option Explicit

Public Enum InstructionProcessorStatus
    ipsIdle = 0
    ipsPlayingNote = 1
    ipsPlayingSilence = 2
    ipsExecutingCommand = 3
End Enum

Public Enum ExecuteSubActions
    esaPlay = 0
    esaPause = 1
    esaStop = 2
End Enum

Private moclInstructionProcessorGlobalVars As Collection
Private mlInstructionProcessorGlobalVarColUserCount As Long

Public Const glINSTRUCTION_PLAY_NOTE As Long = 1
Public Const glINSTRUCTION_SET_INSTRUMENT As Long = 2
Public Const glINSTRUCTION_SET_OCTAVE As Long = 3
Public Const glINSTRUCTION_SET_VOLUME As Long = 4
Public Const glINSTRUCTION_SET_PAN As Long = 5
Public Const glINSTRUCTION_PLAY_SILENCE As Long = 6
Public Const glINSTRUCTION_JUMP As Long = 7
Public Const glINSTRUCTION_SET_VAR As Long = 8
Public Const glINSTRUCTION_IF As Long = 9
Public Const glINSTRUCTION_ARITHMATIC As Long = 10
Public Const glINSTRUCTION_SET_NOTE_INTERVAL As Long = 11
Public Const glINSTRUCTION_SET_SILENCE_INTERVAL As Long = 12
Public Const glINSTRUCTION_PLAY_SUB As Long = 13
Public Const glINSTRUCTION_STOP_SUB As Long = 14
Public Const glINSTRUCTION_PAUSE_SUB As Long = 15
Public Const glINSTRUCTION_RELEASE_CHANNEL As Long = 16
Public Const glINSTRUCTION_NOTE_START As Long = 17
Public Const glINSTRUCTION_NOTE_STOP As Long = 18
Public Const glINSTRUCTION_NO_OPERATION As Long = 19
Public Const glINSTRUCTION_TEMPO As Long = 20
Public Const glINSTRUCTION_RANDOM As Long = 21
Public Const glINSTRUCTION_PRINT As Long = 22


Public Const glARRAY_INDEX_INSTRUCTION_CODE As Long = 0
Public Const glARRAY_INDEX_FIRST_PARAM As Long = 1

Public Const glARITHMATIC_PARAM_VAR_NAME As Long = 1
Public Const glARITHMATIC_PARAM_OPERATION As Long = 2
Public Const glARITHMATIC_PARAM_OPERAND1 As Long = 3
Public Const glARITHMATIC_PARAM_OPERAND2 As Long = 4

Public Const glPLAY_NOTE_PARAM_NOTE As Long = 1
Public Const glPLAY_NOTE_PARAM_NOTE_INTERVAL As Long = 2
Public Const glPLAY_NOTE_PARAM_SILENCE_INTERVAL As Long = 3
Public Const glPLAY_NOTE_PARAM_OCTAVE As Long = 4
Public Const glPLAY_NOTE_PARAM_VOLUME As Long = 5
'Public Const glPLAY_NOTE_PARAM_PAN As Long = 6
'Public Const glPLAY_NOTE_PARAM_INSTRUMENT As Long = 7

Public Const glIF_PARAM_VAR_NAME As Long = 1
Public Const glIF_PARAM_CONDITION As Long = 2
Public Const glIF_PARAM_VALUE As Long = 3
Public Const glIF_PARAM_ON_TRUE_JUMP As Long = 4
Public Const glIF_PARAM_ON_FALSE_JUMP As Long = 5

Public Const glSET_VAR_PARAM_VAR_NAME As Long = 1
Public Const glSET_VAR_PARAM_VAR_VALUE As Long = 2

Public Const glPLAY_SUB_PARAM_SUB_NAME As Long = 1
Public Const glPLAY_SUB_PARAM_LABLE_NAME As Long = 2

Public Const glNOTE_START_STOP_PARAM_NOTE_NUMBER As Long = 1
Public Const glNOTE_START_STOP_PARAM_OCTAVE As Long = 2

Public Const glRANDOM_PARAM_VAR_NAME As Long = 1
Public Const glRANDOM_PARAM_VAR_UPPER_LIMIT As Long = 2

Public Const glPRINT_PARAM_VAR_OR_PARAM_NAME As Long = 1
Public Const glPRINT_PARAM_TAG_WORD As Long = 2



Public Function GetInstructionProcessorGlobalVarCol() As Collection
    
    If moclInstructionProcessorGlobalVars Is Nothing Then
    
        mlInstructionProcessorGlobalVarColUserCount = 1
        
        Set moclInstructionProcessorGlobalVars = New Collection
    
    Else
    
        mlInstructionProcessorGlobalVarColUserCount = mlInstructionProcessorGlobalVarColUserCount + 1
    
    End If
    
    Set GetInstructionProcessorGlobalVarCol = moclInstructionProcessorGlobalVars
    
End Function

Public Sub ReleaseInstructionProcessorGlobalVarCol()

    mlInstructionProcessorGlobalVarColUserCount = mlInstructionProcessorGlobalVarColUserCount - 1
    
    If mlInstructionProcessorGlobalVarColUserCount <= 0 Then
            
        Set moclInstructionProcessorGlobalVars = Nothing
    
    End If
        
End Sub

Public Sub SetInstructionProcessorGlobalVar(ByVal vsVarName As String, ByVal vvVarValue As Variant, Optional ByVal vBoolOnlyIfNew As Boolean = False)
    
    On Error Resume Next
    
    If vBoolOnlyIfNew Then
        
        Dim vTemp As Variant
        
        vTemp = moclInstructionProcessorGlobalVars(vsVarName)
        
        If err.Number <> 0 Then
        
            Call moclInstructionProcessorGlobalVars.Add(vvVarValue, vsVarName)
            
        End If
    
    Else
    
        Call moclInstructionProcessorGlobalVars.Remove(vsVarName)
        
        Call moclInstructionProcessorGlobalVars.Add(vvVarValue, vsVarName)
    
    End If
    
End Sub

Public Function GetInstructionProcessorGlobalVarValue(ByVal vsVarName As String) As Variant
    On Error GoTo ERR_GetInstructionProcessorGlobalVarValue
    GetInstructionProcessorGlobalVarValue = moclInstructionProcessorGlobalVars(vsVarName)
Exit Function
ERR_GetInstructionProcessorGlobalVarValue:
    If err.Number = 5 Then
        err.Clear
        err.Raise lERR_RUN_INVALID_VAR_NAME, , "Non existant variable " & vsVarName
    Else
        ReRaiseError
    End If
End Function

