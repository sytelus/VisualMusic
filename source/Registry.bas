Attribute VB_Name = "modRegistry"
Option Explicit

'This code is by Richard Gardner (http://www.rgsoftware.com/software.htm).
'Some reformating done by Shital Shah.

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const ERROR_SUCCESS = 0&

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal Hkey As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal Hkey As Long, ByVal lpValueName As String) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
    Public Const REG_SZ = 1
    Public Const REG_DWORD = 4


Dim r As Long

Private Declare Function fCreateShellLink Lib "STKIT432.DLL" (ByVal _
    lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal _
    lpstrLinkPath As String, ByVal lpstrLinkArgs As String) As Long


Public Sub SaveKey(ByVal Hkey As Long, ByVal strPath As String)
    Dim keyhand&
    r = RegCreateKey(Hkey, strPath, keyhand&)
    r = RegCloseKey(keyhand&)
End Sub

Public Function GetString(ByVal Hkey As Long, ByVal strPath As String, ByVal strValue As String) As String

    Dim keyhand As Long
    Dim datatype As Long
    Dim lResult As Long
    Dim strBuf As String
    Dim lDataBufSize As Long
    Dim intZeroPos As Integer
    Dim lValueType As Long
    
    r = RegOpenKey(Hkey, strPath, keyhand)
    lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)

    If lValueType = REG_SZ Then
        strBuf = String(lDataBufSize, " ")
        lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)

        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))

            If intZeroPos > 0 Then
                GetString = Left$(strBuf, intZeroPos - 1)
            Else
                GetString = strBuf
            End If

        End If

    End If

End Function

Public Sub SaveString(ByVal Hkey As Long, ByVal strPath As String, ByVal strValue As String, ByVal strdata As String)
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(Hkey, strPath, keyhand)
    r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
    r = RegCloseKey(keyhand)
End Sub

Function GetDWord(ByVal Hkey As Long, ByVal strPath As String, ByVal strValueName As String) As Long

    Dim lResult As Long
    Dim lValueType As Long
    Dim lBuf As Long
    Dim lDataBufSize As Long
    Dim r As Long
    Dim keyhand As Long
    r = RegOpenKey(Hkey, strPath, keyhand)
    lDataBufSize = 4
    lResult = RegQueryValueEx(keyhand, strValueName, 0&, lValueType, lBuf, lDataBufSize)
    If lResult = ERROR_SUCCESS Then
        If lValueType = REG_DWORD Then
            GetDWord = lBuf
        End If
    End If
    r = RegCloseKey(keyhand)
End Function

Public Sub SaveDWord(ByVal Hkey As Long, ByVal strPath As String, ByVal strValueName As String, ByVal lData As Long)
    Dim lResult As Long
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(Hkey, strPath, keyhand)
    lResult = RegSetValueEx(keyhand, strValueName, 0&, REG_DWORD, lData, 4)
    r = RegCloseKey(keyhand)
End Sub

Public Sub DeleteKey(ByVal Hkey As Long, ByVal strKey As String)
    Dim r As Long
    r = RegDeleteKey(Hkey, strKey)
End Sub

Public Sub DeleteValue(ByVal Hkey As Long, ByVal strPath As String, ByVal strValue As String)
    Dim keyhand As Long
    r = RegOpenKey(Hkey, strPath, keyhand)
    r = RegDeleteValue(keyhand, strValue)
    r = RegCloseKey(keyhand)
End Sub

Public Sub AddToRecentDocumentsMenu(ByVal sFileName As String)
    Dim strString As String
    Dim lngDword As Long
    Dim lReturn As Long
    lReturn = fCreateShellLink("..\..\Recent", sFileName, sFileName, "")
End Sub

Public Sub RegisterFileExtensionAndIcon(ByVal vsExtensionWithoutDot As String, ByVal vsFileTypeDescription As String, ByVal vsEXEFile As String, Optional ByVal vsIconFile As String = vbNullString)
    'create an entry in the class key
    Call SaveString(HKEY_CLASSES_ROOT, "\." & vsExtensionWithoutDot, "", vsExtensionWithoutDot & "file")
    'content type
    Call SaveString(HKEY_CLASSES_ROOT, "\." & vsExtensionWithoutDot, "Content Type", "text/plain")
    'name
    Call SaveString(HKEY_CLASSES_ROOT, "\" & vsExtensionWithoutDot & "file", "", vsFileTypeDescription)
    'edit flags
    Call SaveDWord(HKEY_CLASSES_ROOT, "\" & vsExtensionWithoutDot & "file", "EditFlags", "0000")
    'file's icon (can be an icon file, or an icon located within a dl
    '     l file)
    Call SaveString(HKEY_CLASSES_ROOT, "\" & vsExtensionWithoutDot & "file\DefaultIcon", "", vsIconFile)
    'Shell
    Call SaveString(HKEY_CLASSES_ROOT, "\" & vsExtensionWithoutDot & "file\Shell", "", "")
    'Shell Open
    Call SaveString(HKEY_CLASSES_ROOT, "\" & vsExtensionWithoutDot & "file\Shell\Open", "", "")
    'Shell open command
    Call SaveString(HKEY_CLASSES_ROOT, "\" & vsExtensionWithoutDot & "file\Shell\Open\command", "", vsEXEFile & " %1")
End Sub





