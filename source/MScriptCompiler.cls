VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MScriptCompiler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const sMODULE_NAME As String = "MScriptCompiler"

Private moInstrumentManager As InstrumentManager
Private moVariableNames As Collection

Public Property Get InstrumentManager() As InstrumentManager
    Set InstrumentManager = moInstrumentManager
End Property

Public Property Set InstrumentManager(ByVal voInstrumentManager As InstrumentManager)
    Set moInstrumentManager = voInstrumentManager
End Property

Public Function Compile(ByVal vsInstructions As String, Optional ByVal vlCompileStartPoint As Long = 1, Optional ByVal vlCompileEndPoint As Long = -1) As Instructions
    
    On Error GoTo Err_Compile
    
    Dim oclInstructions As Instructions
    Dim sInstruction As String
    Dim vaCompiledInstruction As Variant
    Dim lIndexInInstructionsText As Long
    Dim lLastIndexInInstructionsText As Long
    
    Set oclInstructions = New Instructions
    Set moVariableNames = New Collection
    
    lIndexInInstructionsText = vlCompileStartPoint
    
    sInstruction = GetNextInstruction(vsInstructions, lIndexInInstructionsText)
    lLastIndexInInstructionsText = 0
    
    Dim bExitLoop As Boolean
    bExitLoop = False
    
    Do While (sInstruction <> vbNullString) And (Not bExitLoop)
        vaCompiledInstruction = CompileAInstruction(sInstruction, oclInstructions.Count + 1, oclInstructions.Lables)
        If Not IsEmpty(vaCompiledInstruction) Then
            Call oclInstructions.Add(vaCompiledInstruction)
        End If
        If vlCompileEndPoint <> -1 Then
            If lIndexInInstructionsText >= vlCompileEndPoint Then
                bExitLoop = True
            End If
        End If
        If Not bExitLoop Then
            lLastIndexInInstructionsText = lIndexInInstructionsText
            sInstruction = GetNextInstruction(vsInstructions, lIndexInInstructionsText)
        End If
    Loop
    
    Set Compile = oclInstructions
    Set moVariableNames = Nothing
Exit Function
Err_Compile:
    Set moVariableNames = Nothing
    If ((err.Number >= lERR_BASE_COMPILE) And (err.Number <= lERR_COMPILE_HIGHEST)) Or _
        (err.Number = lERR_COMPILE_NO_LABEL_NAME) Or _
        (err.Number = lERR_COMPILE_INVALID_INSTRUCTION) Then
        err.Raise err.Number, sMODULE_NAME & ":" & lLastIndexInInstructionsText, "Can not compile instruction '" & AlternateStrIfNull(sInstruction, "Instruction text not available>") & "' : " & err.Description
    Else
        ReRaiseError
    End If
End Function

Public Function GetNextInstruction(ByVal vsInstructions As String, ByRef rlIndexInInstructionsText As Long) As String
    
    'First cut of leading white space if any
    Dim lNextSearchPoint As Long
    Dim lInstructionLen As Long
    Dim bInstructionFound As Boolean
    Dim lNextNonWhiteSpaceCharIndex As Long
    Dim lNextWhiteSpaceCharIndex As Long
    
    lInstructionLen = Len(vsInstructions)
    GetNextInstruction = vbNullString
    lNextSearchPoint = rlIndexInInstructionsText
    bInstructionFound = False
    
    Do While (Not bInstructionFound) And (lNextSearchPoint <= lInstructionLen)
            
        'Get the start of next word
        lNextNonWhiteSpaceCharIndex = FindNextNonWhiteSpace(vsInstructions, lNextSearchPoint)
        lNextSearchPoint = lNextNonWhiteSpaceCharIndex
        'e O:4 B:2,,++1 s:11 D#:2,,++3 s:5 F#:2,,++2 s:5 F:2,,++1 s:12 E:2, s:12 B:2,,++1 s:11 F:2,,++3 s:5 G#:2,,++4 s:6 F#:2,,++2 s:12 A:2, s:13 B:2,,++1 s:11 D:2,,++3 s:5 F#:1,,++2 s:6 G#:3,,++1 s:11 B:2, s:11 B:2,,++2 s:11 G#:2,,++2 s:6 C:1,,++2 s:6 E:2,,++1 s:25 A#:2,,++3 s:12 F:2,,++2 s:6 C:1,,++2 s:4 E:3,,++1 s:11 A:1,,++2 s:15 A:1,,++1 s:12 A:2,,++2 s:5 F:2,,++2 s:5 C:2,,++2 s:12 B:2, s:11 C:2,,++2 s:12 D:2,,++3 s:5 B:2,,++2 s:5 G:2,,++2 s:12 C:2,,++3 s:12 E:3,,++2 s:10 C:2,,++2 s:11 G:2,,++2
            
        'If word is found
        If lNextSearchPoint <> -1 Then
            
            Dim sFirstChar As String
            Dim sFirstTwoChars As String
            sFirstChar = Mid$(vsInstructions, lNextSearchPoint, 1)
            sFirstTwoChars = Mid$(vsInstructions, lNextSearchPoint, 2)
            
            'Check if it's comment
            If sFirstChar = "'" Or sFirstTwoChars = "--" Then
                
                'Find next vbCR or vbLF
                Dim lSavedlNextSearchPointValue As Long
                
                lSavedlNextSearchPointValue = lNextSearchPoint
                lNextSearchPoint = InStr(lNextSearchPoint, vsInstructions, vbCr)
                If lNextSearchPoint = 0 Then
                    lNextSearchPoint = InStr(lSavedlNextSearchPointValue, vsInstructions, vbLf)
                End If
                
                'No next line char found
                If lNextSearchPoint = 0 Then
                    lNextSearchPoint = lInstructionLen + 1
                    GetNextInstruction = vbNullString
                End If
                
            Else
                
                'Find the end of instruction and set the bInstructionFound =True
                lNextSearchPoint = FindNextWhiteSpace(vsInstructions, lNextSearchPoint)
                If lNextSearchPoint <> -1 Then
                    GetNextInstruction = Mid$(vsInstructions, lNextNonWhiteSpaceCharIndex, lNextSearchPoint - lNextNonWhiteSpaceCharIndex)
                    rlIndexInInstructionsText = lNextSearchPoint
                Else
                    GetNextInstruction = Mid$(vsInstructions, lNextNonWhiteSpaceCharIndex)
                    rlIndexInInstructionsText = Len(vsInstructions) + 1
                End If
                
                bInstructionFound = True
                
            End If
            
        Else
        
            'No word found at all. Simply exit with null string as return
            GetNextInstruction = vbNullString
            rlIndexInInstructionsText = Len(vsInstructions) + 1
            bInstructionFound = True
            
        End If
        
    Loop

End Function

Public Function CompileAInstruction(ByVal vsInstruction As String, ByVal vlCurrentInstructionCount As Long, ByVal voclLables As Collection) As Variant
    
    Dim lColonPos As Long
    Dim lNextCommaPos As Long
    Dim sInstructionCode As String
    Dim vaInstructionInfo As Variant
    Dim sInstructionParam As String
    Dim lNoteNumber As Long
    
    'Find col in instruction
    lColonPos = InStr(1, vsInstruction, ":")
    
    If lColonPos <> 0 Then
        sInstructionCode = Mid$(vsInstruction, 1, lColonPos - 1)
        sInstructionParam = Trim$(Mid$(vsInstruction, lColonPos + 1))
    Else
        sInstructionCode = vsInstruction
        sInstructionParam = vbNullString
    End If
    
    sInstructionCode = LCase$(sInstructionCode)
    vaInstructionInfo = Empty
    ReDim vaInstructionInfo(0 To 0)
    
    Select Case sInstructionCode
        Case "c": Call BuildInstructionInfoForNote(0, sInstructionParam, vaInstructionInfo)
        Case "c#": Call BuildInstructionInfoForNote(1, sInstructionParam, vaInstructionInfo)
        Case "d": Call BuildInstructionInfoForNote(2, sInstructionParam, vaInstructionInfo)
        Case "d#": Call BuildInstructionInfoForNote(3, sInstructionParam, vaInstructionInfo)
        Case "e": Call BuildInstructionInfoForNote(4, sInstructionParam, vaInstructionInfo)
        Case "f": Call BuildInstructionInfoForNote(5, sInstructionParam, vaInstructionInfo)
        Case "f#": Call BuildInstructionInfoForNote(6, sInstructionParam, vaInstructionInfo)
        Case "g": Call BuildInstructionInfoForNote(7, sInstructionParam, vaInstructionInfo)
        Case "g#": Call BuildInstructionInfoForNote(8, sInstructionParam, vaInstructionInfo)
        Case "a": Call BuildInstructionInfoForNote(9, sInstructionParam, vaInstructionInfo)
        Case "a#": Call BuildInstructionInfoForNote(10, sInstructionParam, vaInstructionInfo)
        Case "b": Call BuildInstructionInfoForNote(11, sInstructionParam, vaInstructionInfo)
        Case "i", "instrument":
            Call BuildInstructionInfo(glINSTRUCTION_SET_INSTRUMENT, sInstructionParam, vaInstructionInfo, 1, 1)
            Dim sInstrumentName As String
            Dim lInstrumentNumber As Long
            If Not IsValidNumber(vaInstructionInfo(1)) Then
                sInstrumentName = Replace(vaInstructionInfo(1), "_", " ", , , vbTextCompare)
                If Not (moInstrumentManager Is Nothing) Then
                    lInstrumentNumber = moInstrumentManager.GetInstrumentNumber(sInstrumentName, -1)
                    If lInstrumentNumber <> -1 Then
                        vaInstructionInfo(1) = lInstrumentNumber
                    ElseIf Not IsItemExistInCol(moVariableNames, vaInstructionInfo(1)) Then
                        err.Raise lERR_COMPILE_INVALID_INSTRUMENT, , "Invalid Instrument name '" & AlternateStrIfNull(sInstrumentName, "<No instrument name supplied>") & "'"
                    End If
                Else
                    err.Raise lERR_COMPILE_NO_INSTRUMENT_NAME_SUPPORT, , "Compile time instrument name resolution not available. This may be because the file '" & gsDEFAULT_INSTRUMENTS_INI & "' does not exist. You can also use instrument number instead of instrument name"
                End If
            End If
        Case "o", "octave":
            Call BuildInstructionInfo(glINSTRUCTION_SET_OCTAVE, sInstructionParam, vaInstructionInfo, 1, 1)
        Case "pan":
            Call BuildInstructionInfo(glINSTRUCTION_SET_PAN, sInstructionParam, vaInstructionInfo, 1, 1)
        Case "v", "volume":
            Call BuildInstructionInfo(glINSTRUCTION_SET_VOLUME, sInstructionParam, vaInstructionInfo, 1, 1)
        Case "p":
            Call BuildInstructionInfo(glINSTRUCTION_SET_PAN, sInstructionParam, vaInstructionInfo, 1, 1)
        Case "j", "jmp", "jump", "goto":
            Call BuildInstructionInfo(glINSTRUCTION_JUMP, sInstructionParam, vaInstructionInfo, 1, 1)
        Case "l", "label":
            If sInstructionParam <> vbNullString Then
                If IsItemExistInCol(voclLables, sInstructionParam) Then
                    err.Raise lERR_COMPILE_LABEL_ALREADY_DECLARED, , "Label '" & sInstructionParam & "' already declared at instruction #" & voclLables(sInstructionParam)
                Else
                    Call voclLables.Add(vlCurrentInstructionCount, sInstructionParam)
                    vaInstructionInfo = Empty
                End If
            Else
                err.Raise lERR_COMPILE_NO_LABEL_NAME, , "No lable name specified"
            End If
        Case "s", "silence", "wait_for", "waitfor":
            Call BuildInstructionInfo(glINSTRUCTION_PLAY_SILENCE, sInstructionParam, vaInstructionInfo, 1, 1)
        Case "set", "var", "dim":
            Call BuildInstructionInfo(glINSTRUCTION_SET_VAR, sInstructionParam, vaInstructionInfo, 1, 2)
            If Not IsItemExistInCol(moVariableNames, vaInstructionInfo(1)) Then
                Call moVariableNames.Add(vaInstructionInfo(1), vaInstructionInfo(1))
            End If
        Case "if":
            Call BuildInstructionInfo(glINSTRUCTION_IF, sInstructionParam, vaInstructionInfo, 4, 5)
        Case "m", "math", "maths":
            Call BuildInstructionInfo(glINSTRUCTION_ARITHMATIC, sInstructionParam, vaInstructionInfo, 2, 4)
        Case "ni", "noteinterval", "note_interval":
            Call BuildInstructionInfo(glINSTRUCTION_SET_NOTE_INTERVAL, sInstructionParam, vaInstructionInfo, 1, 1)
        Case "si", "silenceinterval", "silence_interval":
            Call BuildInstructionInfo(glINSTRUCTION_SET_SILENCE_INTERVAL, sInstructionParam, vaInstructionInfo, 1, 1)
        Case "n", "note":
            Call BuildInstructionInfoForNoteAsNumber(glINSTRUCTION_PLAY_NOTE, sInstructionParam, vaInstructionInfo, 1, 5)
        Case "play", "p", "gosub"
            Call BuildInstructionInfo(glINSTRUCTION_PLAY_SUB, sInstructionParam, vaInstructionInfo, 1, 2)
        Case "stop", "exit"
            Call BuildInstructionInfo(glINSTRUCTION_STOP_SUB, sInstructionParam, vaInstructionInfo, 0, 2)
        Case "pause"
            Call BuildInstructionInfo(glINSTRUCTION_PAUSE_SUB, sInstructionParam, vaInstructionInfo, 0, 2)
        Case "release"
            Call BuildInstructionInfo(glINSTRUCTION_RELEASE_CHANNEL, sInstructionParam, vaInstructionInfo, 0, 0)
        Case "non", "noteon", "note_on"
            Call BuildInstructionInfo(glINSTRUCTION_NOTE_START, sInstructionParam, vaInstructionInfo, 1, 1)
            lNoteNumber = GetNoteNumberForNoteName(vaInstructionInfo(1))
            If lNoteNumber <> -1 Then
                vaInstructionInfo(1) = lNoteNumber
            End If
        Case "noff", "noteoff", "note_off"
            Call BuildInstructionInfo(glINSTRUCTION_NOTE_STOP, sInstructionParam, vaInstructionInfo, 1, 2)
            lNoteNumber = GetNoteNumberForNoteName(vaInstructionInfo(1))
            If lNoteNumber <> -1 Then
                vaInstructionInfo(1) = lNoteNumber
            End If
        Case "nop"
            Call BuildInstructionInfo(glINSTRUCTION_SET_NOTE_INTERVAL, sInstructionParam, vaInstructionInfo, 0, 0)
        Case "tempo"
            Call BuildInstructionInfo(glINSTRUCTION_TEMPO, sInstructionParam, vaInstructionInfo, 1, 1)
        Case "random"
            Call BuildInstructionInfo(glINSTRUCTION_RANDOM, sInstructionParam, vaInstructionInfo, 1, 2)
        Case "print"
            Call BuildInstructionInfo(glINSTRUCTION_PRINT, sInstructionParam, vaInstructionInfo, 0, 2)
        Case Else
            err.Raise lERR_COMPILE_INVALID_INSTRUCTION, , "Unrecognised instruction"
    End Select
    
    CompileAInstruction = vaInstructionInfo
    
End Function

Private Function GetNoteNumberForNoteName(ByVal vsNoteName As String) As Long
    Dim lNoteNumber As Long
    Select Case LCase$(vsNoteName)
        Case "c": lNoteNumber = 0
        Case "c#": lNoteNumber = 1
        Case "d": lNoteNumber = 2
        Case "d#": lNoteNumber = 3
        Case "e": lNoteNumber = 4
        Case "f": lNoteNumber = 5
        Case "f#": lNoteNumber = 6
        Case "g": lNoteNumber = 7
        Case "g#": lNoteNumber = 8
        Case "a": lNoteNumber = 9
        Case "a#": lNoteNumber = 10
        Case "b": lNoteNumber = 11
        Case Else
            lNoteNumber = -1
    End Select
    GetNoteNumberForNoteName = lNoteNumber
End Function

Public Sub BuildInstructionInfoForNote(ByVal vlNoteNumber As Long, ByVal vsNoteParam As String, ByRef rvaInstructionInfo As Variant)
    ReDim Preserve rvaInstructionInfo(0 To 1)
    rvaInstructionInfo(0) = glINSTRUCTION_PLAY_NOTE
    rvaInstructionInfo(1) = vlNoteNumber
    Call StuffParamInArray(vsNoteParam, rvaInstructionInfo)
End Sub

Public Sub BuildInstructionInfoForNoteAsNumber(ByVal vlInstructionCode As Long, ByVal vsNoteParam As String, ByRef rvaInstructionInfo As Variant, Optional ByVal vlMinParamExpected As Variant = Empty, Optional ByVal vlMaxParamExpected As Variant = Empty)
    ReDim Preserve rvaInstructionInfo(0 To 0)
    rvaInstructionInfo(0) = vlInstructionCode
    Call StuffParamInArray(vsNoteParam, rvaInstructionInfo, vlMinParamExpected, vlMaxParamExpected)
End Sub

Public Sub StuffParamInArray(ByVal vsIbstructionParam As String, ByRef vvaInstructionInfo As Variant, Optional ByVal vlMinParamExpected As Variant = Empty, Optional ByVal vlMaxParamExpected As Variant = Empty)
    
    If vsIbstructionParam <> vbNullString Then
        Dim lCommaPos As Long
        Dim lScanIndex As Long
        Dim sParam As String
        Dim lInstructionParamLen As String
        Dim lInstructionInforArrayIndex As Long
        Dim lParamCount As Long
        
        lInstructionInforArrayIndex = UBound(vvaInstructionInfo)
        
        lInstructionParamLen = Len(vsIbstructionParam)
        
        lScanIndex = 1
        
        Do While lScanIndex <= lInstructionParamLen
        
            lCommaPos = InStr(lScanIndex, vsIbstructionParam, ",")
            
            If lCommaPos <> 0 Then
                sParam = Mid$(vsIbstructionParam, lScanIndex, lCommaPos - lScanIndex)
                lScanIndex = lCommaPos + 1
            Else
                sParam = Mid$(vsIbstructionParam, lScanIndex)
                lScanIndex = lInstructionParamLen + 1
            End If
                   
            sParam = Trim$(sParam)
               
            lInstructionInforArrayIndex = lInstructionInforArrayIndex + 1
        
            ReDim Preserve vvaInstructionInfo(0 To lInstructionInforArrayIndex)
            lParamCount = lParamCount + 1
            
            If sParam <> vbNullString Then
                vvaInstructionInfo(lInstructionInforArrayIndex) = sParam
            Else
                vvaInstructionInfo(lInstructionInforArrayIndex) = Empty
            End If
        
        Loop
        
    End If
    
    If Not IsEmpty(vlMinParamExpected) Then
        If lParamCount < vlMinParamExpected Then
            err.Raise lERR_COMPILE_MISSING_PARAM, , "Required parameter(s) not specified"
        End If
    End If
    If Not IsEmpty(vlMaxParamExpected) Then
        If lParamCount > vlMaxParamExpected Then
            err.Raise lERR_COMPILE_MORE_PARAM, , "More parameter(s) then required specified"
        End If
    End If
    
End Sub

Private Sub BuildInstructionInfo(ByVal vlInstructionCode As Long, ByVal vsInstructionParam As String, ByRef rvaInstructionInfo As Variant, Optional ByVal vlMinParamExpected As Variant = Empty, Optional ByVal vlMaxParamExpected As Variant = Empty)
    rvaInstructionInfo(0) = vlInstructionCode
    Call StuffParamInArray(vsInstructionParam, rvaInstructionInfo, vlMinParamExpected, vlMaxParamExpected)
End Sub

Private Sub Class_Initialize()
    Call InitInstrumentList
End Sub

Private Sub Class_Terminate()
    Set moInstrumentManager = Nothing
End Sub

Private Sub InitInstrumentList()
    On Error GoTo ERR_InitInstrumentList
    Set moInstrumentManager = New InstrumentManager
    moInstrumentManager.INIFileName = gsDEFAULT_INSTRUMENTS_INI
Exit Sub
ERR_InitInstrumentList:
    Set moInstrumentManager = Nothing
End Sub
