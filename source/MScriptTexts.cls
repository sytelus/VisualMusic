VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MScripts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Private variable to hold collection
Private mCol As Collection
Private moclMScriptNames As Collection
Private moMScriptCompiler As MScriptCompiler
Private moInstructionClock As ClockProvider
Private mlTimerInterval As Long
Private msLastCompiledMScript As String

Public Event AddMScript(ByVal vsMScriptName As String, ByVal vsMScriptText As String)
Public Event ModifyMScriptText(ByVal vsMScriptName As String, ByVal vsOldMScriptText As String, ByVal vsNewMScriptText As String, ByRef rbCancel As Boolean)
Public Event StartRenameMScript(ByVal vsOldMScriptName As String, ByVal vsNewMScriptName As String, ByRef rbCanceled As Boolean)
Public Event EndRenameMScript(ByVal vsOldMScriptName As String, ByVal vsNewMScriptName As String)
Public Event DeleteMScript(ByVal vsMScriptName As String, ByRef rbCancel As Boolean)
Public Event DeleteAllMScript(ByRef rbCancel As Boolean)
Public Event InstructionProcessorStateChanged(ByVal vsMScriptName As String, ByVal venmOldState As InstructionProcessorStatus, ByVal venmNewState As InstructionProcessorStatus)
Public Event InstructionProcessorError(ByVal vsMScriptName As String, ByVal Number As Long, ByVal Source As String, ByVal Description As String, ByVal HelpFile As String, ByVal HelpContext As Long)
Public Event InstructionProcessorExecutionStartStopEvent(ByVal vsMScriptName As String, ByVal vboolStartStopFlag As Boolean)
Public Event PrintRequest(ByVal vsMScriptName As String, ByVal vsStringToPrint As String)

Private Const sMSCRIPT_SIGNATURE As String = "MScript"
Private Const sMSCRIPT_TIMER_INTERVAL_INSTRUCTION As String = "TIMER"
Private Const sMSCRIPT_VER_INSTRUCTION As String = "VER"
Private Const lMSCRIPT_THIS_VER As Long = 1
Private Const lMSCRIPT_COMPATIBLE_VER As Long = 1
Private Const sTUNE_NAME_START As String = "TuneStart"
Private Const sTUNE_NAME_END As String = "TuneEnd"
Private Const lDEFAULT_TIMER_INTERVAL As Long = 60

Property Get LastCompiledMScript() As String
    LastCompiledMScript = msLastCompiledMScript
End Property

Friend Property Let LastCompiledMScript(ByVal vsMScriptName As String)
     msLastCompiledMScript = vsMScriptName
End Property

Public Property Get TimerInterval() As Long
   TimerInterval = mlTimerInterval
End Property

Public Property Get InstructionClock() As ClockProvider
    Set InstructionClock = moInstructionClock
End Property

Public Property Set InstructionClock(voInstructionClock As ClockProvider)
    Set moInstructionClock = voInstructionClock
End Property

Public Property Get LOWER_BOUND() As Long

    'Tell the caller that index value for this collection starts from 1
    LOWER_BOUND = 1

End Property

Public Property Get Count() As Long

    'Simply return the count stored in private collection
    Count = mCol.Count

End Property

Public Property Get Item(ByVal Index As Variant) As MScript
Attribute Item.VB_UserMemId = 0
    
    'Error could be generated because
    'specified index doesn't exist.
    On Error GoTo ERR_Get_Item

    'Store the refrence stored in private collection
    Set Item = mCol.Item(Index)
    
Exit Property
ERR_Get_Item:

    Call HandleColError(Index)

End Property

'This function gives the ability to the collection
'for using For Each statement with it

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

    'Return the private collection's
    'hidden enum member
    Set NewEnum = mCol.[_NewEnum]

End Property

Public Function Clear()

    Dim bCanceled As Boolean
    
    bCanceled = False
    
    RaiseEvent DeleteAllMScript(bCanceled)
    
    If Not bCanceled Then
    
        Call ClearCollection(mCol)
        Call ClearCollection(moclMScriptNames)

    Else
    
        err.Raise 1000, , "Operation canceled by event handler"

    End If

End Function

Public Function IsExists(ByVal Name As String) As Boolean

    'Use delayed error handling
    On Error Resume Next
    
    Dim vItem As Variant
    
    'Try to access specified item
    Set vItem = mCol.Item(Name)
    
    'If there is no error then specified item exist
    If err.Number = 0 Then
    
        'then return True
        IsExists = True
    
    'If error is "Invalid procedure call or argument", it
    'means specified item does not exist in collection
    ElseIf err.Number = 5 Then
        
        'else return false
        IsExists = False
    
    'Some other unexpected error has occured
    Else
    
        'So reraise it
        ReRaiseError
    
    End If

End Function

Public Function Add(ByVal vsMScriptText As String, ByVal vsMScriptName As String) As MScript

    'Error could be generated because
    'specified field name may exist already
    On Error GoTo ERR_Add

    Dim oMScript As MScript

    Set oMScript = New MScript
    
    oMScript.Name = vsMScriptName
    Set oMScript.ParentMScriptsCol = Me
    oMScript.Text = vsMScriptText
    Set oMScript.InstructionClock = moInstructionClock
    Set oMScript.MScriptCompiler = moMScriptCompiler

    'Add the field refrence in the collection with the field name as the key
    Call mCol.Add(oMScript, vsMScriptName)
    Call moclMScriptNames.Add(vsMScriptName, vsMScriptName)
    
    RaiseEvent AddMScript(vsMScriptName, vsMScriptText)
    
    Set Add = oMScript
            
Exit Function
ERR_Add:

    'Some error has occured while adding the
    'item to the collection. The HandleColError
    'function converts this error to more
    'specific error for this custom collection
    HandleColError

End Function

Public Sub Remove(ByVal vsMScriptName As String)

    'Error could be generated because
    'specified index doesn't exist.
    On Error GoTo ERR_Remove

    Dim bCanceled As Boolean
    bCanceled = False

    RaiseEvent DeleteMScript(vsMScriptName, bCanceled)

    If Not bCanceled Then
    
        'Remove the specified item from the collection
        Call mCol.Remove(vsMScriptName)
        Call moclMScriptNames.Remove(vsMScriptName)
    
    Else
    
        err.Raise 1000, , "Operation canceled by event handler"

    End If
    
Exit Sub
ERR_Remove:

    'Some error has occured while getting the
    'removing the item from the collection.
    'The HandleColError function converts this
    'error to more specific error for this custom
    'collection.
    HandleColError

End Sub

Private Sub Class_Initialize()
    
    'create the private collection
    Set mCol = New Collection
    Set moclMScriptNames = New Collection
    mlTimerInterval = lDEFAULT_TIMER_INTERVAL
End Sub

Public Property Get MScriptCompiler() As MScriptCompiler
    Set MScriptCompiler = moMScriptCompiler
End Property

Public Property Set MScriptCompiler(voMScriptCompiler As MScriptCompiler)
    Set moMScriptCompiler = voMScriptCompiler
End Property

Private Sub Class_Terminate()

    'Remove all items from the collection
    Call Me.Clear
        
    'Destroy the private collection
    Set moMScriptCompiler = Nothing
    Set moInstructionClock = Nothing
    Set moclMScriptNames = Nothing
    Set mCol = Nothing
End Sub

Private Sub HandleColError(Optional ByVal vvAditionalInfo As Variant = Empty)

    'Check the error number
    Select Case err.Number

        'Err 5: Invalid Procedure call or argument - occures when wrong key is supplied to collection
        Case 5
        
            err.Raise 1000, , "Specified MScript tune " & AlternateStrIfNull(vvAditionalInfo, "<Tune name not available>") & " does not exist."
            
        'Err 9: Subscript out of range - occures when invalid index number is specified
        Case 9
            
            err.Raise 1000, , "There is no MScript tune for the specified index"
            
        'Err 457: This Key is already associtaed with an element of this collection
        Case 457
        
            err.Raise 1000, , "A MScript tune with the same name already present in the collection."
        
        'Unexpected error has occured
        Case Else

            'So just reraise it
            ReRaiseError
            
    End Select

End Sub

Public Function SaveToText() As String
'    Dim sFullScript As String
'    Dim sMScript As String
'    Dim sMScriptName As String
'    Dim i As Integer
'    Const sDELIM_STR As String = vbCrLf & "*****" & vbCrLf
'
'    sFullScript = vbNullString
'
'    For i = 1 To moclMScriptNames.Count
'        sMScriptName = moclMScriptNames(i)
'        sFullScript = sFullScript & sDELIM_STR & sMScriptName & sDELIM_STR
'        sFullScript = sFullScript & mCol(sMScriptName).Text
'    Next i
'    SaveToText = sFullScript
    
    Dim sScriptText As String
    Dim sTuneName As String
    Dim sTuneText As String
    Dim lTuneIndex As Long
    
    'Add signature
    sScriptText = sMSCRIPT_SIGNATURE & " "
    
    'Add version compatibility info
    sScriptText = sScriptText & sMSCRIPT_VER_INSTRUCTION & ":" & lMSCRIPT_THIS_VER & vbCrLf
    
    'Add timer interval info
    sScriptText = sScriptText & sMSCRIPT_TIMER_INTERVAL_INSTRUCTION & ":" & mlTimerInterval & vbCrLf & vbCrLf

    For lTuneIndex = 1 To moclMScriptNames.Count
        sTuneName = moclMScriptNames(lTuneIndex)
        sTuneText = mCol(sTuneName).Text
        
        sScriptText = sScriptText & sTUNE_NAME_START & ":" & sTuneName & vbCrLf
        sScriptText = sScriptText & sTuneText
        sScriptText = sScriptText & vbCrLf & sTUNE_NAME_END & ":" & sTuneName & vbCrLf
        
    Next lTuneIndex
    
    SaveToText = sScriptText
    
End Function

Public Sub LoadFromText(ByVal vsMScripts As String, Optional ByVal vboolCheckSignature As Boolean = True, Optional ByVal vboolCheckVersion As Boolean = True)
'    Dim sFullScript As String
'    Dim lStart As Long
'    Dim lStop As Long
'    Dim lTuneEnd As Long
'    Dim sTuneName As String
'    Dim sTuneText As String
'    Dim lTuneCount As Long
'    Const sDELIM_STR As String = vbCrLf & "*****" & vbCrLf
'
'    lStart = 1
'    lTuneCount = 1
'    Call Me.Clear
'    sFullScript = vsMScripts
'    If sFullScript <> vbNullString Then
'        Do
'            lStart = InStr(lStart, sFullScript, sDELIM_STR, vbBinaryCompare) + Len(sDELIM_STR)
'            lStop = InStr(lStart, sFullScript, sDELIM_STR, vbBinaryCompare)
'            sTuneName = Mid$(sFullScript, lStart, lStop - lStart)
'            lTuneEnd = InStr(lStop + Len(sDELIM_STR), sFullScript, sDELIM_STR, vbBinaryCompare)
'            If lTuneEnd <> 0 Then
'                sTuneText = Mid$(sFullScript, lStop + Len(sDELIM_STR), lTuneEnd - lStop - Len(sDELIM_STR))
'            Else
'                sTuneText = Mid$(sFullScript, lStop + Len(sDELIM_STR))
'            End If
'            Call Me.Add(sTuneText, sTuneName)
'            lStart = lTuneEnd
'            lTuneCount = lTuneCount + 1
'        Loop While lTuneEnd <> 0
'    End If

    Call Me.Clear

    Dim lCurPos As Long
    Dim sWord As String
    Dim vaArrayForWord As Variant
    
    'Check the signature
    lCurPos = 1
    sWord = GetNextWord(vsMScripts, lCurPos, lCurPos)
    vaArrayForWord = ConvertColonCommaWordToArray(sWord)
    If vboolCheckSignature Then
        If vaArrayForWord(LBound(vaArrayForWord)) <> sMSCRIPT_SIGNATURE Then
            err.Raise 1000, , "Not valid MScript: MScript signature not found"
        End If
    End If
    
    'Check Version
    sWord = GetNextWord(vsMScripts, lCurPos, lCurPos)
    If vboolCheckVersion Then
        Dim vaCompatibleVersions As Variant
        vaCompatibleVersions = ConvertColonCommaWordToArray(sWord)
        If vaCompatibleVersions(LBound(vaCompatibleVersions)) <> sMSCRIPT_VER_INSTRUCTION Then
            err.Raise 1000, , "Not Valid MScript: Information for compatible versions not found"
        Else
            Dim lVersionArrayIndex As Long
            Dim bCompatibleVersionFound As Boolean
            bCompatibleVersionFound = False
            For lVersionArrayIndex = LBound(vaCompatibleVersions) + 1 To UBound(vaCompatibleVersions)
                If IsValidNumber(VarType(vaCompatibleVersions(lVersionArrayIndex))) Then
                    If lMSCRIPT_COMPATIBLE_VER >= vaCompatibleVersions(lVersionArrayIndex) Then
                        bCompatibleVersionFound = True
                        Exit For
                    End If
                Else
                    If LCase$(vaCompatibleVersions(lVersionArrayIndex)) = "vmusic1" Then
                        bCompatibleVersionFound = True
                        Exit For
                    End If
                End If
            Next lVersionArrayIndex
            If Not bCompatibleVersionFound Then
                err.Raise 1000, , "MScript is not compatible with current player's version " & lMSCRIPT_COMPATIBLE_VER
            End If
        End If
    End If
    
    'Get timer interval info
    sWord = GetNextWord(vsMScripts, lCurPos, lCurPos)
    vaArrayForWord = ConvertColonCommaWordToArray(sWord)
    If GetDimension(vaArrayForWord) < 2 Then
        err.Raise 1000, , "MScript not valid: Expected timer interval information not found"
    Else
        If UCase$(vaArrayForWord(LBound(vaArrayForWord))) <> sMSCRIPT_TIMER_INTERVAL_INSTRUCTION Then
            err.Raise 1000, , "MScript not valid: No timer interval instruction found"
        Else
            If Not IsValidNumber(vaArrayForWord(LBound(vaArrayForWord) + 1)) Then
                err.Raise 1000, , "MScript not valid: Timer interval '" & AlternateStrIfNull(vaArrayForWord(LBound(vaArrayForWord) + 1), "<empty string>") & "' is not a valid number"
            Else
                mlTimerInterval = vaArrayForWord(LBound(vaArrayForWord) + 1)
            End If
        End If
    End If
    
    'Get the tune texts
    Dim sTuneName As String
    Dim sTuneText As String
    Dim lTuneIndex As Long
    Dim lTotalTextLen As Long
    Dim lTuneTextStart As Long
    
    lTotalTextLen = Len(vsMScripts)
    
    Do While lTotalTextLen >= lCurPos
        'Find the tune start
        lCurPos = InStr(lCurPos, vsMScripts, sTUNE_NAME_START, vbTextCompare)
        If lCurPos <> 0 Then
            sWord = GetNextWord(vsMScripts, lCurPos, lCurPos)
            If sWord <> vbNullString Then
                lTuneTextStart = lCurPos + 1
                vaArrayForWord = ConvertColonCommaWordToArray(sWord)
                If (UBound(vaArrayForWord) - LBound(vaArrayForWord) + 1) < 2 Then
                    err.Raise 1000, , "Invalid MScript: " & sTUNE_NAME_START & " doesn't have tune name part at position " & lCurPos
                End If
                sTuneName = vaArrayForWord(LBound(vaArrayForWord) + 1)
                lCurPos = InStr(lCurPos, vsMScripts, vbCrLf & sTUNE_NAME_END & ":" & sTuneName, vbTextCompare)
                If lCurPos <> 0 Then
                    sTuneText = Mid$(vsMScripts, lTuneTextStart, lCurPos - lTuneTextStart)
                Else
                    err.Raise 1000, , "Invalid MScript: End of tune " & sTuneName & " not found"
                End If
                
                Call Me.Add(sTuneText, sTuneName)
            End If
        Else
            lCurPos = lTotalTextLen + 1
        End If
    Loop
    
End Sub

Public Sub ClearAllText()
    Dim lMScriptsIndex As Long
    
    For lMScriptsIndex = 1 To mCol.Count
        Me.Item(lMScriptsIndex).Text = vbNullString
    Next lMScriptsIndex
End Sub

Friend Sub GenerateModifyMScriptEvent(ByVal vsMScriptName As String, ByVal vsOldMScriptText As String, ByVal vsNewMScriptText As String, ByRef rbCancel As Boolean)
    RaiseEvent ModifyMScriptText(vsMScriptName, vsOldMScriptText, vsNewMScriptText, rbCancel)
End Sub
Friend Sub GeneratePrintRequestEvent(ByVal vsMScriptName As String, ByVal vsStringToPrint As String)
    RaiseEvent PrintRequest(vsMScriptName, vsStringToPrint)
End Sub
Friend Sub GenerateInstructionProcessorStateChangedEvent(ByVal vsMScriptName As String, ByVal venmOldState As InstructionProcessorStatus, ByVal venmNewState As InstructionProcessorStatus)
    RaiseEvent InstructionProcessorStateChanged(vsMScriptName, venmOldState, venmNewState)
End Sub

Friend Sub GenerateInstructionProcessorErrorEvent(ByVal vsMScriptName As String, ByVal Number As Long, ByVal Source As String, ByVal Description As String, ByVal HelpFile As String, ByVal HelpContext As Long)
    RaiseEvent InstructionProcessorError(vsMScriptName, Number, Source, Description, HelpFile, HelpContext)
End Sub

Friend Sub GenerateInstructionProcessorExecutionStartStopEvent(ByVal vsMScriptName As String, ByVal vboolStartStopFlag As Boolean)
    RaiseEvent InstructionProcessorExecutionStartStopEvent(vsMScriptName, vboolStartStopFlag)
End Sub

Public Sub RenameMScript(ByVal vsOldMScriptName As String, ByVal vsNewMScriptName As String)
            
    If Not IsExists(vsNewMScriptName) Then
        
        Dim bCanceled As Boolean
        
        bCanceled = False
        
        RaiseEvent StartRenameMScript(vsOldMScriptName, vsNewMScriptName, bCanceled)
        
        If Not bCanceled Then
        
            Dim oMScript As MScript
            
            'Get the object first
            Set oMScript = Item(vsOldMScriptName)
            
            'Reset the MScript object's name
            oMScript.Name = vsNewMScriptName
            
            'Now remove this item
            Call moclMScriptNames.Remove(vsOldMScriptName)
            Call mCol.Remove(vsOldMScriptName)
            
            'Add the object with new name
            Call moclMScriptNames.Add(vsNewMScriptName, vsNewMScriptName)
            Call mCol.Add(oMScript, vsNewMScriptName)
            
            Set oMScript = Nothing
            
            RaiseEvent EndRenameMScript(vsOldMScriptName, vsNewMScriptName)
        
        Else
        
            err.Raise 1000, , "Rename canceled"
            
        End If
    
    Else
    
        err.Raise 1000, , "MScript with name " & vsNewMScriptName & " already exist"
    
    End If
    
End Sub

Public Sub CompileAll()

    Dim oMScript As MScript
    
    For Each oMScript In mCol
        msLastCompiledMScript = oMScript.Name
        oMScript.Compile
    Next oMScript

End Sub

Public Sub PlayAll()

    Dim oMScript As MScript
    
    For Each oMScript In mCol
        msLastCompiledMScript = oMScript.Name
        oMScript.Play
    Next oMScript

End Sub
