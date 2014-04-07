VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InstrumentManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private moclInstrumentNumbers As Collection
Private moclInstrumentNames As Collection
Private m_sINIFileName As String
Private m_bAutoAddAppPath As Boolean

Public Event InstrumentListChanged()
Public Event InstrumentNameChange(ByVal vlInstrumentNumber As Long, ByVal vsOldInstrumentName As String, ByVal vsNewInstrumentName As String, ByVal vsINIFileName As String, ByRef Cancle As Boolean)
Private Const sINI_SECTION As String = "General MIDI"

Public Property Get AutoAddAppPath() As Boolean
    AutoAddAppPath = m_bAutoAddAppPath
End Property

Public Property Let AutoAddAppPath(ByVal vboolAutoAddAppPath As Boolean)
    m_bAutoAddAppPath = vboolAutoAddAppPath
End Property

Public Property Get INIFileName() As String
    INIFileName = m_sINIFileName
End Property

Public Property Let INIFileName(ByVal vsINIFileName As String)
    m_sINIFileName = GetINIFileNameWithAutoPathAdd(vsINIFileName)
    Call LoadInstrumentsFromINI(m_sINIFileName)
End Property

Public Function GetINIFileNameWithAutoPathAdd(ByVal vsINIFileName As String) As String
    If m_bAutoAddAppPath And (Trim$(vsINIFileName) <> vbNullString) Then
        If (InStr(1, vsINIFileName, "\") = 0) And (InStr(1, vsINIFileName, ":") = 0) Then
            GetINIFileNameWithAutoPathAdd = GetPathWithSlash(App.Path) & vsINIFileName
        Else
            GetINIFileNameWithAutoPathAdd = vsINIFileName
        End If
    Else
        GetINIFileNameWithAutoPathAdd = vsINIFileName
    End If
End Function

Private Sub Class_Initialize()
    Set moclInstrumentNumbers = New Collection
    Set moclInstrumentNames = New Collection
    m_bAutoAddAppPath = True
End Sub

Private Sub Class_Terminate()
    Set moclInstrumentNumbers = Nothing
    Set moclInstrumentNames = Nothing
End Sub

Private Sub LoadInstrumentsFromINI(ByVal vsFileName As String)
    Dim sInstrumentName As String
    Dim sKey As String
    Dim lInstrumentNum As Long
    Dim lNumOfCharReturned As Long
    Dim lsiListItem As ListItem
    
    If Not IsFileExist(vsFileName) Then
        err.Raise 1000, , AlternateStrIfNull(vsFileName, "<No file name specified>") & " file containing instrument information doesn't exist"
    End If
    
    If Trim$(vsFileName) <> vbNullString Then
    
        Call ClearCollection(moclInstrumentNames)
        Call ClearCollection(moclInstrumentNumbers)
        
        For lInstrumentNum = 0 To 127
            sKey = Str$(lInstrumentNum)
            sInstrumentName = String$(255, 0)
            lNumOfCharReturned = GetPrivateProfileString(sINI_SECTION, sKey, vbNullString, sInstrumentName, Len(sInstrumentName), vsFileName)
            If lNumOfCharReturned <> 0 Then
                sInstrumentName = Left$(sInstrumentName, lNumOfCharReturned)
                If Trim$(sInstrumentName) <> vbNullString Then
                    Call moclInstrumentNames.Add(sInstrumentName, CStr(lInstrumentNum))
                    Call moclInstrumentNumbers.Add(CStr(lInstrumentNum), sInstrumentName)
                End If
            End If
        Next lInstrumentNum
        
    Else
    
        Call ClearCollection(moclInstrumentNames)
        Call ClearCollection(moclInstrumentNumbers)
        
    End If
    RaiseEvent InstrumentListChanged
End Sub

Public Function GetInstrumentNumber(ByVal vvInstrumentNameOrIndex As Variant, Optional ByVal vvDefaultInstrumentIfInvalidName As Variant) As Long
On Error GoTo ERR_GetInstrumentNumber
    GetInstrumentNumber = moclInstrumentNumbers(vvInstrumentNameOrIndex)
Exit Function
ERR_GetInstrumentNumber:
    If IsMissing(vvDefaultInstrumentIfInvalidName) Then
        ReRaiseError
    Else
        If err.Number = 5 Then  'Specified key doesn't exist in collection
            GetInstrumentNumber = vvDefaultInstrumentIfInvalidName
        Else
            ReRaiseError
        End If
    End If
End Function

Public Function GetInstrumentName(ByVal vlInstrumentNumberOrIndex As Long, Optional ByVal vboolIsIndex As Boolean = False) As String
    If Not vboolIsIndex Then
        GetInstrumentName = moclInstrumentNames(CStr(vlInstrumentNumberOrIndex))
    Else
        GetInstrumentName = moclInstrumentNames(vlInstrumentNumberOrIndex)
    End If
End Function

Public Function IsInstrumentExist(ByVal vvNameOrNumber As Variant)
    If VarType(vvNameOrNumber) = vbString Then
        IsInstrumentExist = IsItemExistInCol(moclInstrumentNames, vvNameOrNumber)
    Else
        IsInstrumentExist = IsItemExistInCol(moclInstrumentNumbers, vvNameOrNumber)
    End If
End Function

Public Sub SetInstrumentName(ByVal vlInstrumentNumber As Long, ByVal vsNewInstrumentName As String, Optional ByVal vsINIFileName As Variant = Empty)
    Dim bCancleChange As Boolean
    Dim sINIFileName As String
    
    If Not IsEmpty(vsINIFileName) Then
        sINIFileName = GetINIFileNameWithAutoPathAdd(vsINIFileName)
    Else
        sINIFileName = m_sINIFileName
    End If
    
    bCancleChange = False
    RaiseEvent InstrumentNameChange(vlInstrumentNumber, GetInstrumentName(vlInstrumentNumber), vsNewInstrumentName, sINIFileName, bCancleChange)
    If Not bCancleChange Then
        Call moclInstrumentNames.Remove(CStr(vlInstrumentNumber))
        Call moclInstrumentNames.Add(vsNewInstrumentName, CStr(vlInstrumentNumber))
        Call WritePrivateProfileString(sINI_SECTION, CStr(vlInstrumentNumber), vsNewInstrumentName, sINIFileName)
    End If
End Sub

Public Function FilterInstrumentList(ByVal vsWordList As String) As Long
    
    Dim oclWordsToSearch As Collection
    Dim oclWordsAvailable As Collection
    Dim bMatchFound As Boolean
    Dim lWordToSearchIndex As Long
    Dim lWordsAvailableIndex As Long
    Dim lInstrumentsIndex As Long
    Dim sWordToSearch As String
    Dim sWordAvailable As String
    
    Set oclWordsToSearch = New Collection
    Set oclWordsAvailable = New Collection
    
    'Build the word list
    Call MakeWordList(vsWordList, oclWordsToSearch)
    
    If oclWordsToSearch.Count > 0 Then
        
        For lInstrumentsIndex = moclInstrumentNames.Count To 1 Step -1
        
            Call MakeWordList(moclInstrumentNames(lInstrumentsIndex), oclWordsAvailable)
            
            bMatchFound = False
            For lWordsAvailableIndex = oclWordsAvailable.Count To 1 Step -1
                sWordAvailable = oclWordsAvailable(lWordsAvailableIndex)
                For lWordToSearchIndex = 1 To oclWordsToSearch.Count
                    sWordToSearch = oclWordsToSearch(lWordToSearchIndex)
                    If InStr(1, sWordAvailable, sWordToSearch, vbTextCompare) <> 0 Then
                        bMatchFound = True
                        Exit For
                    End If
                Next lWordToSearchIndex
                If bMatchFound Then Exit For
            Next lWordsAvailableIndex
        
            If Not bMatchFound Then
                Call moclInstrumentNumbers.Remove(moclInstrumentNames(lInstrumentsIndex))
                Call moclInstrumentNames.Remove(lInstrumentsIndex)
            End If
                    
        Next lInstrumentsIndex
        
    End If
    
    Set oclWordsToSearch = Nothing
    Set oclWordsAvailable = Nothing
    
    RaiseEvent InstrumentListChanged
    
End Function

Public Property Get InstrumentCount()
    InstrumentCount = moclInstrumentNames.Count
End Property

Public Function GetInstrumentType(ByVal vvInstrumentNameOrNumber As ValueConstants) As InstrumentType
    
    Dim lInstrumentNumber As Long
    
    If VarType(vvInstrumentNameOrNumber) = vbString Then
        lInstrumentNumber = GetInstrumentNumber(vvInstrumentNameOrNumber)
    Else
        lInstrumentNumber = vvInstrumentNameOrNumber
    End If
    
    GetInstrumentType = lInstrumentNumber \ 8
    
End Function
