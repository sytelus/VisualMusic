VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.UserControl MScriptEditor 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6405
   EditAtDesignTime=   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   6405
   Begin VB.Frame fraClearButton 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   6000
      TabIndex        =   8
      Top             =   3240
      Width           =   255
      Begin VB.TextBox lblClearMScriptText 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -60
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "X"
         ToolTipText     =   "Clear the script"
         Top             =   60
         Width           =   315
      End
   End
   Begin VB.TextBox txtEditorCover 
      Alignment       =   2  'Center
      BackColor       =   &H00008080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   555
      Left            =   3060
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "MScriptEditor.ctx":0000
      Top             =   60
      Width           =   2355
   End
   Begin VB.PictureBox pctTuneListButtonContainer 
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   750
      ScaleHeight     =   390
      ScaleWidth      =   840
      TabIndex        =   3
      Top             =   3225
      Width           =   840
      Begin VB.TextBox lblDeleteTune 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   540
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "X"
         ToolTipText     =   "Delete Tune (Press Alt and X)"
         Top             =   60
         Width           =   195
      End
      Begin VB.TextBox lblModifyTune 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   240
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   "~"
         ToolTipText     =   "Rename the Tune (Press Alt and ~)"
         Top             =   30
         Width           =   255
      End
      Begin VB.TextBox lblAddTune 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   0
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "+"
         ToolTipText     =   "Add New Tune (Press Alt and +)"
         Top             =   0
         Width           =   195
      End
   End
   Begin VB.PictureBox pctSeperator 
      BorderStyle     =   0  'None
      Height          =   3540
      Left            =   2175
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3540
      ScaleWidth      =   150
      TabIndex        =   2
      Top             =   0
      Width           =   150
   End
   Begin MSComctlLib.ListView lvwTuneList 
      Height          =   3090
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   5450
      View            =   2
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Tune"
         Object.Width           =   2540
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtbTuneScript 
      Height          =   3090
      Left            =   1725
      TabIndex        =   1
      Top             =   150
      Width           =   4740
      _ExtentX        =   8361
      _ExtentY        =   5450
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"MScriptEditor.ctx":0006
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "MScriptEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mbSeperatorDragStarted As Boolean
Private WithEvents m_oMScripts As MScripts
Attribute m_oMScripts.VB_VarHelpID = -1
Private mbIsListViewItemEdited As Boolean

Public Event TuneSelected(ByVal vsTuneName As String)

Public Enum MScriptEditorControls
    mecTuneList = 0
   mecTuneEditor = 1
End Enum

Public Property Get MScripts() As MScripts
    Set MScripts = m_oMScripts
End Property

Public Property Set MScripts(ByVal voMScripts As MScripts)
    Set m_oMScripts = voMScripts
End Property

Public Function ProcessOnKeyDown(ByVal KeyCode As Long, ByVal Shift As Long) As Boolean
    Select Case KeyCode
        Case 187 And ((Shift And vbAltMask) <> 0)   '+/= key
            UserControl.SetFocus
            lvwTuneList.SetFocus
            Call lblAddTune_Click
            ProcessOnKeyDown = True
        Case vbKeyX And ((Shift And vbAltMask) <> 0)
            UserControl.SetFocus
            lvwTuneList.SetFocus
            Call lblDeleteTune_Click
            ProcessOnKeyDown = True
        'For US kbd
        Case 192 And ((Shift And vbAltMask) <> 0) 'Tilde key: ~
            UserControl.SetFocus
            lvwTuneList.SetFocus
            Call lblModifyTune_Click
            ProcessOnKeyDown = True
        'For UK kbd
        Case 222 And ((Shift And vbAltMask) <> 0) 'Tilde key: ~
            UserControl.SetFocus
            lvwTuneList.SetFocus
            Call lblModifyTune_Click
            ProcessOnKeyDown = True
        Case Else
            ProcessOnKeyDown = False
    End Select
End Function

Private Sub lblAddTune_Click()
    Dim sTuneName As String
    sTuneName = InputBox("Enter the name for the new tune:", "Tune Name", vbNullString)
    If sTuneName <> vbNullString Then
        If m_oMScripts.IsExists(sTuneName) Then
            MsgBox "A tune with same name '" & sTuneName & "' already exist"
        Else
            Call m_oMScripts.Add(vbNullString, sTuneName)
            Call rtbTuneScript.SetFocus
        End If
    End If
End Sub

Private Sub lblAddTune_GotFocus()
    lvwTuneList.SetFocus
End Sub

Private Sub lblAddTune_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call SetLabelStyle(lblAddTune, True)
End Sub

Private Sub lblClearMScriptText_Click()
    rtbTuneScript.Text = vbNullString
End Sub

Private Sub lblClearMScriptText_GotFocus()
    If rtbTuneScript.Visible And rtbTuneScript.Enabled Then
        rtbTuneScript.SetFocus
    Else
        lvwTuneList.SetFocus
    End If
End Sub

Private Sub lblClearMScriptText_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call SetLabelStyle(lblClearMScriptText, True)
End Sub

Private Sub lblDeleteTune_Click()
    Dim lviSelectedItem As ListItem
    Set lviSelectedItem = lvwTuneList.SelectedItem
    If Not (lviSelectedItem Is Nothing) Then
        Call m_oMScripts.Remove(ConvertToNormalKey(lviSelectedItem.Text))
    Else
        MsgBox "First select the tune you want to delete"
    End If
End Sub

Private Sub lblDeleteTune_GotFocus()
    lvwTuneList.SetFocus
End Sub

Private Sub lblDeleteTune_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        Call SetLabelStyle(lblDeleteTune, True)
End Sub

Private Sub lblModifyTune_Click()
    Dim lviSelectedItem As ListItem
    Set lviSelectedItem = lvwTuneList.SelectedItem
    If Not (lviSelectedItem Is Nothing) Then
        UserControl.SetFocus
        lvwTuneList.SetFocus
        lvwTuneList.StartLabelEdit
    Else
        MsgBox "First select the tune that you want to rename"
    End If
    
End Sub

Private Sub lblModifyTune_GotFocus()
    lvwTuneList.SetFocus
End Sub

Private Sub lblModifyTune_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call SetLabelStyle(lblModifyTune, True)
End Sub

Private Sub lvwTuneList_AfterLabelEdit(Cancel As Integer, NewString As String)

    On Error GoTo ERR_lvwTuneList_AfterLabelEdit

    If Trim$(NewString) <> vbNullString Then
        mbIsListViewItemEdited = False
        Dim lviSelectedItem As ListItem
        Set lviSelectedItem = lvwTuneList.SelectedItem
        If Not (lviSelectedItem Is Nothing) Then
            Call m_oMScripts.RenameMScript(lviSelectedItem.Text, NewString)
            lviSelectedItem.Key = ConvertToListViewKey(NewString)
            RaiseEvent TuneSelected(NewString)
        Else
            Cancel = True
            err.Raise 1000, , "Select the tune you want to rename first"
        End If
    Else
        Cancel = True
        err.Raise 1000, , "Tune name can not be blank"
    End If
    
Exit Sub
ERR_lvwTuneList_AfterLabelEdit:
    Cancel = True
    Call DisplayErrorMessage
End Sub

Private Sub lvwTuneList_BeforeLabelEdit(Cancel As Integer)
    mbIsListViewItemEdited = True
End Sub

Private Sub lvwTuneList_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call RefreshTuneTextForSelectedTune
End Sub

Private Sub lvwTuneList_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete
            Call lblDeleteTune_Click
        Case vbKeyF2
            lvwTuneList.StartLabelEdit
            Call lblModifyTune_Click
    End Select
End Sub

Private Sub lvwTuneList_LostFocus()
    Call SaveCurrentText
End Sub

Private Sub lvwTuneList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call SaveCurrentText
End Sub

Private Sub lvwTuneList_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call ResetLabelStyle
End Sub

Private Sub m_oMScripts_AddMScript(ByVal vsMScriptName As String, ByVal vsMScriptText As String)
    Set lvwTuneList.SelectedItem = lvwTuneList.ListItems.Add(, ConvertToListViewKey(vsMScriptName), vsMScriptName)
    rtbTuneScript.Text = vbNullString
    Call RefreshEditorCover
End Sub

Private Sub m_oMScripts_DeleteAllMScript(rbCancel As Boolean)
    Call lvwTuneList.ListItems.Clear
    rtbTuneScript.Text = vbNullString
    Call RefreshEditorCover
End Sub

Private Sub m_oMScripts_DeleteMScript(ByVal vsMScriptName As String, rbCancel As Boolean)
    Call lvwTuneList.ListItems.Remove(ConvertToListViewKey(vsMScriptName))
    Call RefreshTuneTextForSelectedTune
    Call RefreshEditorCover
End Sub

Private Sub m_oMScripts_ModifyMScriptText(ByVal vsMScriptName As String, ByVal vsOldMScriptText As String, ByVal vsNewMScriptText As String, rbCancel As Boolean)
    Dim lviSelectedItem As ListItem
    Set lviSelectedItem = lvwTuneList.SelectedItem
    If Not (lviSelectedItem Is Nothing) Then
        If lviSelectedItem.Text = vsMScriptName Then
            rtbTuneScript.Text = vsNewMScriptText
        End If
    End If
End Sub

'Private Sub m_oMScripts_EndRenameMScript(ByVal vsOldMScriptName As String, ByVal vsNewMScriptName As String)
'    Call lvwTuneList.ListItems.Remove(vsOldMScriptName)
'    Call lvwTuneList.ListItems.Add(, vsNewMScriptName, vsNewMScriptName)
'    lvwTuneList.ListItems(vsNewMScriptName).Selected = True
'    Call RefreshTuneTextForSelectedTune
'End Sub

Private Sub pctSeperator_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mbSeperatorDragStarted = True
    Call ReleaseCapture
    Call SetCapture(hwnd)
End Sub

Private Sub pctTuneListButtonContainer_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call ResetLabelStyle
End Sub

Private Sub rtbTuneScript_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 Then
        
        If KeyCode = vbKeyTab Then
            Dim sFirstPart As String
            Dim sSecondPart As String
            Dim sText As String
            Dim lSelStart As Long
            
            lSelStart = rtbTuneScript.SelStart
            
            sText = rtbTuneScript.Text
            
            If lSelStart <> 0 Then
                sFirstPart = Mid(sText, 1, lSelStart)
            Else
                sFirstPart = vbNullString
            End If
            
            If lSelStart >= Len(sText) Then
                sSecondPart = vbNullString
            Else
                sSecondPart = Mid$(sText, lSelStart + 1, Len(sText) - lSelStart)
            End If
            
            rtbTuneScript.Text = sFirstPart & vbTab & sSecondPart
            rtbTuneScript.SelStart = lSelStart + 1
            KeyCode = 0
        End If
    End If
End Sub

Private Sub rtbTuneScript_LostFocus()
    Call SaveCurrentText
End Sub

Private Sub rtbTuneScript_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call ResetLabelStyle
End Sub

Private Sub UserControl_EnterFocus()
    Call ResetLabelStyle
End Sub

Private Sub UserControl_ExitFocus()
    Call SaveCurrentText
    Call ResetLabelStyle
End Sub

Private Sub UserControl_Initialize()
    mbIsListViewItemEdited = False
End Sub

Private Sub UserControl_LostFocus()
    Call SaveCurrentText
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If mbSeperatorDragStarted Then
        pctSeperator.Left = x
    End If
    Call ResetLabelStyle
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    mbSeperatorDragStarted = False
    Call ReleaseCapture
    If pctSeperator.Left < 0 Then
        pctSeperator.Left = pctSeperator.Width
    ElseIf pctSeperator.Left > Width Then
        pctSeperator.Left = Width - pctSeperator.Width
    End If
    UserControl_Resize
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Call RefreshEditorCover
End Sub

Private Sub UserControl_Resize()
    
    Const lLABLE_SPACING As Long = 30
    
    pctSeperator.Width = 35
    
    lvwTuneList.Left = 0
    lvwTuneList.Top = 0
    lvwTuneList.Height = Height - lblAddTune.Height - lLABLE_SPACING
    lvwTuneList.Width = pctSeperator.Left - lvwTuneList.Left
    
'    lblAddTune.Top = lvwTuneList.Left + lvwTuneList.Height + CLng(lLABLE_SPACING / 2)
'    lblModifyTune.Top = lblAddTune.Top
'    lblDeleteTune.Top = lblAddTune.Top

    pctTuneListButtonContainer.Top = lvwTuneList.Left + lvwTuneList.Height + CLng(lLABLE_SPACING / 2)
    pctTuneListButtonContainer.Left = lvwTuneList.Left + lvwTuneList.Width - pctTuneListButtonContainer.Width
        
    pctSeperator.Left = lvwTuneList.Left + lvwTuneList.Width
    pctSeperator.Top = lvwTuneList.Top
    pctSeperator.Height = lvwTuneList.Height
    
    rtbTuneScript.Left = pctSeperator.Left + pctSeperator.Width
    rtbTuneScript.Top = 0
    rtbTuneScript.Height = lvwTuneList.Height
    rtbTuneScript.Width = Width - ((lvwTuneList.Left * 2) + lvwTuneList.Width + pctSeperator.Width)
    
    fraClearButton.Left = rtbTuneScript.Left + rtbTuneScript.Width - fraClearButton.Width
    fraClearButton.Top = pctTuneListButtonContainer.Top
    
End Sub

Public Sub RefreshTuneTextForSelectedTune()
    Dim lviSelectedItem As ListItem
    Set lviSelectedItem = lvwTuneList.SelectedItem
    If Not (lviSelectedItem Is Nothing) Then
        rtbTuneScript.Text = m_oMScripts(lviSelectedItem.Text).Text
        RaiseEvent TuneSelected(lviSelectedItem.Text)
    Else
        rtbTuneScript.Text = vbNullString
        RaiseEvent TuneSelected(vbNullString)
    End If
End Sub

Public Sub SaveCurrentText(Optional ByVal vsCurrentTuneName As Variant)
    
    Dim sCurrentTuneName As String
    
    Dim lviSelectedItem As ListItem
    Set lviSelectedItem = lvwTuneList.SelectedItem
    
    If Not IsMissing(vsCurrentTuneName) Then
        sCurrentTuneName = vsCurrentTuneName
    Else
        If Not (lviSelectedItem Is Nothing) Then
            sCurrentTuneName = lviSelectedItem.Text
        Else
            sCurrentTuneName = vbNullString
        End If
    End If
    
    If Not (sCurrentTuneName = vbNullString) Then
        m_oMScripts(sCurrentTuneName).Text = rtbTuneScript.Text
    End If
    
End Sub

Public Function PlaySelectedTune() As String
    Call SaveCurrentText
    Dim lviSelectedItem As ListItem
    Set lviSelectedItem = lvwTuneList.SelectedItem
    If Not (lviSelectedItem Is Nothing) Then
        If rtbTuneScript.SelLength = 0 Then
            m_oMScripts(lviSelectedItem.Text).Play
        Else
            Call m_oMScripts(lviSelectedItem.Text).Play(rtbTuneScript.SelStart, rtbTuneScript.SelStart + rtbTuneScript.SelLength - 1)
        End If
        PlaySelectedTune = lviSelectedItem.Text
    Else
        PlaySelectedTune = vbNullString
    End If
End Function

Public Function GetSelectedTune() As String
    Dim lviSelectedItem As ListItem
    Set lviSelectedItem = lvwTuneList.SelectedItem
    If Not (lviSelectedItem Is Nothing) Then
        GetSelectedTune = lviSelectedItem.Text
    Else
        GetSelectedTune = vbNullString
    End If
End Function

Public Function CompileSelectedTune() As String
    Call SaveCurrentText
    Dim lviSelectedItem As ListItem
    Set lviSelectedItem = lvwTuneList.SelectedItem
    If Not (lviSelectedItem Is Nothing) Then
        Call m_oMScripts(lviSelectedItem.Text).Compile
        CompileSelectedTune = lviSelectedItem.Text
    Else
        CompileSelectedTune = vbNullString
    End If
End Function

Public Property Get Text() As String
    Text = rtbTuneScript.Text
End Property

Public Property Let Text(ByVal vsMScriptText As String)
    rtbTuneScript.Text = vsMScriptText
End Property

Public Property Get EnableEdit() As Boolean
    EnableEdit = rtbTuneScript.Enabled
End Property

Public Property Let EnableEdit(ByVal vboolEnableEdit As Boolean)
    rtbTuneScript.Enabled = vboolEnableEdit
End Property

Public Sub SelectTune(ByVal vsTuneNameOrIndex As Variant, Optional ByVal vboolFailSafe As Boolean = False)
    On Error GoTo ERR_SelectTune
    If VarType(vsTuneNameOrIndex) = vbString Then
        lvwTuneList.ListItems(ConvertToListViewKey(vsTuneNameOrIndex)).Selected = True
    Else
        lvwTuneList.ListItems(vsTuneNameOrIndex).Selected = True
    End If
    Call RefreshTuneTextForSelectedTune
Exit Sub
ERR_SelectTune:
    If Not vboolFailSafe Then
        ReRaiseError
    End If
End Sub

Public Sub MoveTextCursor(ByVal vlPosition As Long)
    rtbTuneScript.SelStart = vlPosition
    rtbTuneScript.SetFocus
End Sub

Public Function FocusOnEditor() As Boolean
    FocusOnEditor = (ActiveControl Is rtbTuneScript)
End Function

Public Function IsListViewItemEdited() As Boolean
    IsListViewItemEdited = mbIsListViewItemEdited
End Function

Public Sub SetFocusToControl(ByVal venmControl As MScriptEditorControls)
    Select Case venmControl
        Case mecTuneEditor
            rtbTuneScript.SetFocus
        Case mecTuneList
            lvwTuneList.SetFocus
        Case Else
            err.Raise 1000, , "Code to set focus to new control not written"
    End Select
End Sub

Public Sub RefreshEditorCover()
    Dim lviSelectedItem As ListItem
    Set lviSelectedItem = lvwTuneList.SelectedItem
    If Not (lviSelectedItem Is Nothing) Then
        txtEditorCover.Visible = False
        RaiseEvent TuneSelected(lviSelectedItem.Text)
    Else
        RaiseEvent TuneSelected(vbNullString)
        txtEditorCover.Left = rtbTuneScript.Left
        txtEditorCover.Width = rtbTuneScript.Width
        txtEditorCover.Height = rtbTuneScript.Height
        txtEditorCover.Top = rtbTuneScript.Top
        txtEditorCover.Text = vbCrLf & "(No tunes selected)" & vbCrLf & "Click '+' at the bottom to add new tune"
        txtEditorCover.BackColor = &H4477EE
        txtEditorCover.Visible = True
    End If
End Sub

Private Function SetLabelStyle(ByRef rlbl As TextBox, ByVal vIsMouseOver As Boolean)

    If vIsMouseOver Then
    
        Call ResetLabelStyle
    
        rlbl.ForeColor = vbRed
    
    Else
     
        rlbl.ForeColor = vbButtonText
     
    End If

End Function

Private Function ResetLabelStyle()
    Call SetLabelStyle(lblAddTune, False)
    Call SetLabelStyle(lblModifyTune, False)
    Call SetLabelStyle(lblDeleteTune, False)
    Call SetLabelStyle(lblClearMScriptText, False)
End Function

Private Function ConvertToListViewKey(ByVal vsNormalKey As String) As String
    
    Dim sFirstChar As String
    
    sFirstChar = Left$(vsNormalKey, 1)
    
    If (sFirstChar >= "0" And sFirstChar <= "9") Or sFirstChar = "_" Then
        ConvertToListViewKey = "_" & vsNormalKey
    Else
        ConvertToListViewKey = vsNormalKey
    End If
    
End Function

Private Function ConvertToNormalKey(ByVal vsListViewKey As String) As String
    Dim sFirstChar As String
    
    sFirstChar = Left$(vsListViewKey, 1)
    
    If sFirstChar = "_" Then
        ConvertToNormalKey = Mid$(vsListViewKey, 2)
    Else
        ConvertToNormalKey = vsListViewKey
    End If
End Function

Public Sub SelectWordAt(ByVal vlCharPos As Long)
    
    Dim lNextWordPos As Long
    
    Dim lStartPoint As Long
    
    If rtbTuneScript.SelLength = 0 Then
        lStartPoint = 0
    Else
        lStartPoint = rtbTuneScript.SelStart
    End If
    
    lNextWordPos = FindNextNonWhiteSpace(rtbTuneScript.Text, lStartPoint + vlCharPos)
    
    If lNextWordPos <> -1 Then
        rtbTuneScript.SelStart = lStartPoint + lNextWordPos - 1
        lNextWordPos = FindNextWhiteSpace(rtbTuneScript.Text, lNextWordPos)
        If lNextWordPos <> -1 Then
            rtbTuneScript.SelLength = lNextWordPos - rtbTuneScript.SelStart
        Else
            rtbTuneScript.SelLength = Len(rtbTuneScript.Text) - rtbTuneScript.SelStart
        End If
    End If
    
End Sub
