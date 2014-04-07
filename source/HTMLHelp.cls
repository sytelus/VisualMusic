VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHTMLHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hWndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const HH_DISPLAY_TOPIC As Long = 0
Private Const HH_DISPLAY_TOC As Long = 1
Private Const HH_DISPLAY_INDEX As Long = 1
Private Const HH_DISPLAY_SEARCH As Long = 3
Private Const SW_SHOW = 5

Private m_lParentHWnd As Long
Private m_sHelpFile

Public Property Get ParentHWnd() As Long
    ParentHWnd = m_lParentHWnd
End Property

Public Property Let ParentHWnd(ByVal vlHWnd As Long)
    m_lParentHWnd = vlHWnd
End Property

Public Property Get HelpFile() As String
    HelpFile = m_sHelpFile
End Property

Public Property Let HelpFile(ByVal vsHelpFile As String)
    m_sHelpFile = vsHelpFile
End Property

Public Sub ShowTopic(ByVal vsTopicName As String)
    Dim lHTMLAPIReturn As Long
    lHTMLAPIReturn = HtmlHelp(m_lParentHWnd, m_sHelpFile & "::/" & vsTopicName, HH_DISPLAY_TOPIC, 0)
End Sub

Public Sub SetToDefaults()
    m_lParentHWnd = GetDesktopWindow
    m_sHelpFile = App.HelpFile
End Sub

Public Sub ShowContent()
    Dim lHTMLAPIReturn As Long
    'This doesn't work on other machines
    'lHTMLAPIReturn = HtmlHelp(m_lParentHWnd, m_sHelpFile, HH_DISPLAY_TOC, 0)
    'So now replaced by this
    If Not OpenAnyFile(m_sHelpFile) Then
        MsgBox "Can not open the file " & m_sHelpFile & vbCrLf & "Please try to find this file in the Visual Music directory and double click on it."
    End If
End Sub

Public Sub ShowIndex()
    Dim lHTMLAPIReturn As Long
    lHTMLAPIReturn = HtmlHelp(m_lParentHWnd, m_sHelpFile, HH_DISPLAY_INDEX, 0)
End Sub

Public Sub ShowSearch()
    Dim lHTMLAPIReturn As Long
    lHTMLAPIReturn = HtmlHelp(m_lParentHWnd, m_sHelpFile, HH_DISPLAY_SEARCH, 0)
End Sub

Public Sub GotoWebpage(ByVal vsURL As String)
    Call OpenAnyFile(vsURL)
End Sub

Public Sub InvokeEmailClient(ByVal vsTOAddress As String)
    Call OpenAnyFile("mailto:" & vsTOAddress)
End Sub
Private Function OpenAnyFile(ByVal vsFileName As String, Optional ByVal vsParameters As String = "") As Boolean
    OpenAnyFile = ShellExecute(0, "open", vsFileName, vsParameters, "", SW_SHOW) > 32
End Function

Private Sub Class_Initialize()
    Call SetToDefaults
End Sub
