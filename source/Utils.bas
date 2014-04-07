Attribute VB_Name = "Utils"
Option Explicit

Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOW = 5

Public Sub ClearCollection(ByVal voclCol As Collection)

    Dim lColIndex As Long
    
    For lColIndex = voclCol.Count To 1 Step -1
    
        voclCol.Remove lColIndex
    
    Next lColIndex

End Sub

Public Sub ReRaiseError()
    If err.Number <> 0 Then
        err.Raise err.Number, err.Source, err.Description, err.HelpFile, err.HelpContext
    End If
End Sub

Public Function IsValidNumber(ByVal vvValueToCheck As Variant) As Boolean
    If IsNumeric(vvValueToCheck) Then
        If Not IsEmpty(vvValueToCheck) And Not IsNull(vvValueToCheck) Then
            IsValidNumber = True
        Else
            IsValidNumber = False
        End If
    Else
        IsValidNumber = False
    End If
End Function

Public Function AlternateStrIfNull(ByVal vvntVariable As Variant, ByVal vsAlternateValueOnNull As String) As String

    Dim bReturnAlternate As Boolean
    
    bReturnAlternate = True
    
    If VarType(vvntVariable) = vbObject Then
        
        If Not (vvntVariable Is Nothing) Then
        
            bReturnAlternate = False
            
        Else
        
            bReturnAlternate = True
            
        End If
        
    Else
    
        If Not (IsEmpty(vvntVariable)) Then
        
            If Not (IsNull(vvntVariable)) Then
            
                If (Len(CStr(vvntVariable)) <> 0) Then
                
                    bReturnAlternate = False
                
                End If
                
            End If
            
        End If
        
    End If
    
    
    If bReturnAlternate Then
    
        AlternateStrIfNull = vsAlternateValueOnNull
        
    Else
    
        AlternateStrIfNull = CStr(vvntVariable)
    
    End If

End Function

Public Function IsItemExistInCol(ByVal voCol As Object, ByVal vvntIndexOrKey As String)

    'Enable delayed error handling
    On Error Resume Next

    Dim vnt As Variant      'Temporarily holds col item
    
    'Try to access specified item
    vnt = voCol.Item(vvntIndexOrKey)
    
    'If there is Item property returns object and there is no default
    'property, Err 438 occures (Method or property not in object)
    If err.Number = 438 Then
    
        'Clear the previous error
        err.Clear
    
        'Try to set object in to variant
        Set vnt = voCol.Item(vvntIndexOrKey)
        
    End If
    
    'If error is occured then item doesn't exist
    'TODO: add the exact error codes
    If err.Number = 0 Then
        
        'Item exist in collection, so returns True
        IsItemExistInCol = True
        
    'If error is "Invalid procedure call or argument", it
    'means specified item does not exist in collection
    ElseIf err.Number = 5 Then
    
        'Item doesn't exist in collection, so returns False
        IsItemExistInCol = False
        
    'Some other unexpected error has occured
    Else
    
        'So reraise it
        ReRaiseError
    
    End If

    'Clear the Err object
    err.Clear

End Function

Public Function GetPathWithSlash(ByVal vsPath As String) As String

    'Routine specific local vars here

    'common variables
    Dim sErrorLocation As String            'Location of the error for global err handler
    Dim nOldMousePointer As Integer         'Current State of the mouse pointer.

    'Set the mouse pointer to hour glass
    nOldMousePointer = SetMousePointer(vbHourglass)

    'set the name of the function to be pass for the error traping function
    sErrorLocation = "Utils.GetPathWithSlash"


    If Trim$(vsPath) <> "" Then
    
        If Right$(vsPath, 1) <> "\" Then
            
            GetPathWithSlash = vsPath + "\"
            
        Else
        
            GetPathWithSlash = vsPath
        
        End If

    Else
        
        GetPathWithSlash = ""
        
    End If

    'Set the mouse pointer to prev. state
    nOldMousePointer = SetMousePointer(nOldMousePointer)

End Function

Public Function SetMousePointer(ByVal nCursorType As Integer) As Integer
    Dim nOldMousePointer As Integer
    
    ' get the old mouse pointer
    nOldMousePointer = Screen.MousePointer
    
    ' set the mouse pointer
    Screen.MousePointer = nCursorType
    
    ' return the old mouse pointer
    SetMousePointer = nOldMousePointer
    
End Function '

Public Function FindNextNonWhiteSpace(ByVal vsString As String, ByVal vlStartPoint As Long) As Long
    Dim lNextNonWhiteSpaceIndex As Long
    Dim lStringScanIndex As Long
    Dim sCharToTest As String
    Dim lStartPoint As Long
    
    lNextNonWhiteSpaceIndex = -1
    
    If vlStartPoint > 0 Then
        lStartPoint = vlStartPoint
    Else
        lStartPoint = 1
    End If
    
    For lStringScanIndex = lStartPoint To Len(vsString)
        sCharToTest = Mid$(vsString, lStringScanIndex, 1)
        If Not IsWhiteSpace(sCharToTest) Then
            lNextNonWhiteSpaceIndex = lStringScanIndex
            Exit For
        End If
    Next lStringScanIndex
    
    FindNextNonWhiteSpace = lNextNonWhiteSpaceIndex
End Function

Public Function FindNextWhiteSpace(ByVal vsString As String, ByVal vlStartPoint As Long) As Long
    Dim lNextWhiteSpaceIndex As Long
    Dim lStringScanIndex As Long
    Dim sCharToTest As String
    
    lNextWhiteSpaceIndex = -1
    
    For lStringScanIndex = vlStartPoint To Len(vsString)
        sCharToTest = Mid$(vsString, lStringScanIndex, 1)
        If IsWhiteSpace(sCharToTest) Then
            lNextWhiteSpaceIndex = lStringScanIndex
            Exit For
        End If
    Next lStringScanIndex
    
    FindNextWhiteSpace = lNextWhiteSpaceIndex
End Function

Public Sub MakeWordList(ByVal vsString As String, ByVal voclWords As Collection)
    
    Dim lCurrentSearchPos As Long
    Dim lNextSearchPos As Long
    Dim sWord As String
    
    Call ClearCollection(voclWords)
    
    lCurrentSearchPos = 1
    lCurrentSearchPos = FindNextNonWhiteSpace(vsString, lCurrentSearchPos)
    
    If lCurrentSearchPos <> -1 Then
        Do
            lNextSearchPos = FindNextWhiteSpace(vsString, lCurrentSearchPos)
            If lNextSearchPos <> -1 Then
                sWord = Mid$(vsString, lCurrentSearchPos, lNextSearchPos - lCurrentSearchPos)
                lCurrentSearchPos = FindNextNonWhiteSpace(vsString, lNextSearchPos)
            Else
                sWord = Mid$(vsString, lCurrentSearchPos)
            End If
            Call voclWords.Add(sWord)
        Loop While lNextSearchPos <> -1
    End If
End Sub

Public Function IsWhiteSpace(ByVal vsChar As String) As Boolean
    If vsChar = " " Or vsChar = Chr$(9) _
        Or vsChar = Chr$(13) Or vsChar = Chr$(10) Then
        IsWhiteSpace = True
    Else
        IsWhiteSpace = False
    End If
End Function

Public Function GetNextWord(ByVal vsText As String, ByVal vlStartPoint As Long, Optional ByRef rlNextWordPos As Variant) As String
    
    Dim lWordStart As Long
    Dim lWordEnd As Long
    Dim bEndOfTextReached As Boolean
    Dim sWord As String
    
    lWordStart = FindNextNonWhiteSpace(vsText, vlStartPoint)
    If lWordStart <> -1 Then
        lWordEnd = FindNextWhiteSpace(vsText, lWordStart)
        If lWordEnd <> -1 Then
            sWord = Mid$(vsText, lWordStart, lWordEnd - lWordStart)
            bEndOfTextReached = False
        Else
            sWord = Mid$(vsText, lWordStart)
            bEndOfTextReached = True
        End If
    Else
        bEndOfTextReached = True
        sWord = vbNullString
    End If
    
    If Not IsMissing(rlNextWordPos) Then
        If bEndOfTextReached = True Then
            rlNextWordPos = Len(vsText) + 1
        Else
            rlNextWordPos = lWordEnd + 1
        End If
    End If
    
    GetNextWord = sWord
    
End Function
'Convert a work like "Some_Thing:Param1,Param2,..." in to Array(Some_Thing,Param1,Param2,...)
Public Function ConvertColonCommaWordToArray(ByVal vsWord As String) As Variant
    
    Dim lColonPos As Long
    Dim lNextCommaPos As Long
    Dim sWordBeforeColon As String
    Dim vaArray As Variant
    Dim sWordAfterColon As String
    
    'Find col in instruction
    lColonPos = InStr(1, vsWord, ":")
    
    If lColonPos <> 0 Then
        sWordBeforeColon = Mid$(vsWord, 1, lColonPos - 1)
        sWordAfterColon = Trim$(Mid$(vsWord, lColonPos + 1))
    Else
        sWordBeforeColon = vsWord
        sWordAfterColon = vbNullString
    End If
    
    vaArray = Empty
    ReDim vaArray(0 To 0)
    
    vaArray(0) = sWordBeforeColon
    
    If sWordAfterColon <> vbNullString Then
        Dim lCommaPos As Long
        Dim lScanIndex As Long
        Dim sParam As String
        Dim lWordAfterColonLen As String
        Dim lArrayIndex As Long
        
        lArrayIndex = UBound(vaArray)
        
        lWordAfterColonLen = Len(sWordAfterColon)
        
        lScanIndex = 1
        
        Do While lScanIndex <= lWordAfterColonLen
        
            lCommaPos = InStr(lScanIndex, sWordAfterColon, ",")
            
            If lCommaPos <> 0 Then
                sParam = Mid$(sWordAfterColon, lScanIndex, lCommaPos - lScanIndex)
                lScanIndex = lCommaPos + 1
            Else
                sParam = Mid$(sWordAfterColon, lScanIndex)
                lScanIndex = lWordAfterColonLen + 1
            End If
                   
            sParam = Trim$(sParam)
               
            lArrayIndex = lArrayIndex + 1
        
            ReDim Preserve vaArray(0 To lArrayIndex)
            
            If sParam <> vbNullString Then
                vaArray(lArrayIndex) = sParam
            Else
                vaArray(lArrayIndex) = Empty
            End If
        
        Loop
        
    End If
    
    ConvertColonCommaWordToArray = vaArray
    
End Function

Public Function SaveStringToFile(ByVal vsString As String, ByVal vsFileName As String, Optional ByVal vboolFailIfReadOnlyFile As Boolean = True)
    
    If vboolFailIfReadOnlyFile Then
        If IsFileReadOnly(vsFileName) Then
            err.Raise 1000, , "File is read only"
        End If
    End If
    
    Dim nFileHandle As Integer
    
    'First open file for output mode to trucate it
    nFileHandle = FreeFile
    Open vsFileName For Output Access Write As #nFileHandle
    Close #nFileHandle
    
    'Now write actual data
    nFileHandle = FreeFile
    Open vsFileName For Binary Access Write As #nFileHandle
        Put #nFileHandle, , vsString
    Close #nFileHandle
    
End Function

Public Function LoadStringFromFile(ByVal vsFileName As String)
    Dim nFileHandle As Integer
    Dim sFileContent As String
        
    nFileHandle = FreeFile
    
    Open vsFileName For Binary Access Read As #nFileHandle
        sFileContent = Input(FileLen(vsFileName), #nFileHandle)
    Close #nFileHandle
    LoadStringFromFile = sFileContent
End Function

Public Function IsFileReadOnly(ByVal vsFileName As String) As Boolean
    On Error GoTo ERR_IsFileReadOnly
    IsFileReadOnly = ((GetAttr(vsFileName) And vbReadOnly) <> 0)
Exit Function
ERR_IsFileReadOnly:
    If err.Number = 53 Then 'File diesn't exist
        IsFileReadOnly = False
    Else
        ReRaiseError
    End If
End Function

Public Sub DisplayErrorMessage()
    Dim sDiscription As String
    sDiscription = AlternateStrIfNull(err.Description, "<No error description available>")
    If Right(sDiscription, 1) <> "." Then
        sDiscription = sDiscription & "."
    End If
    If err.Number <> 1000 And err.Number > 0 Then
        MsgBox "Error " & err.Number & " : " & sDiscription
    Else
        MsgBox "Error : " & sDiscription
    End If
End Sub

Public Function IsFileExist(ByVal sFileName As String) As Boolean
    On Error GoTo ERR_IsFileExist
    IsFileExist = Dir$(sFileName, vbNormal Or vbHidden Or vbSystem) <> ""
Exit Function
ERR_IsFileExist:
    IsFileExist = False
End Function
Public Function GetDimension(ByVal varr As Variant) As Integer

    GetDimension = 0
    
    On Error Resume Next
    
    GetDimension = UBound(varr, 1) - LBound(varr, 1) + 1

End Function

Public Function OpenAnyFile(ByVal vsFileName As String, Optional ByVal vsParameters As String = "") As Boolean
    OpenAnyFile = ShellExecute(0, "open", vsFileName, vsParameters, "", SW_SHOW) > 32
End Function

