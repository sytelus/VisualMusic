VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl InstrumentSelector 
   ClientHeight    =   3255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2685
   LockControls    =   -1  'True
   ScaleHeight     =   3255
   ScaleWidth      =   2685
   Begin MSComctlLib.ListView lvwInstruments 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   4683
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDragMode     =   1
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDragMode     =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No."
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Instrument"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblOnOff 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ý"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   300
      TabIndex        =   5
      Top             =   2775
      Width           =   180
   End
   Begin VB.Label lblPlayOnClick 
      AutoSize        =   -1  'True
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   975
      TabIndex        =   4
      Top             =   2700
      Width           =   285
   End
   Begin VB.Line linLablesBottom 
      Visible         =   0   'False
      X1              =   75
      X2              =   2625
      Y1              =   3075
      Y2              =   3075
   End
   Begin VB.Line linLablesTop 
      Visible         =   0   'False
      X1              =   75
      X2              =   2625
      Y1              =   2700
      Y2              =   2700
   End
   Begin VB.Label lblShowHideFav 
      AutoSize        =   -1  'True
      Caption         =   "«"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1275
      TabIndex        =   3
      Top             =   2700
      Width           =   285
   End
   Begin VB.Label lblAddRemoveFav 
      AutoSize        =   -1  'True
      Caption         =   "ü"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1725
      TabIndex        =   2
      Top             =   2850
      Width           =   255
   End
   Begin VB.Label lblFilter 
      AutoSize        =   -1  'True
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   15.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2025
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2625
      Width           =   450
   End
End
Attribute VB_Name = "InstrumentSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event InstrumentSelected(ByVal vlInstrumentNumber As Long, ByVal vsInstrumentName As String)
Public Event PlaySample()
Public Event StopSamplePlay()

Private WithEvents moInstrumentManager As InstrumentManager
Attribute moInstrumentManager.VB_VarHelpID = -1
Private m_sINIFileForInstruments As String
Private m_bAppRelativePath As Boolean
Private m_sINIFileForFavInstruments As String

Private mbFavoritesMode As Boolean
Private mbListIsFiltred As Boolean
Private mbPlayOnClick As Boolean
Private mbSamplePlayStopped As Boolean

Public Property Get SelectedInstrument() As Long
    Dim lsiSelectedList As ListItem
    Set lsiSelectedList = lvwInstruments.SelectedItem
    If Not lsiSelectedList Is Nothing Then
        SelectedInstrument = lsiSelectedList.Text
    Else
        SelectedInstrument = -1
    End If
End Property

Public Property Let SelectedInstrument(ByVal vlInstrumentNumber As Long)
    Dim sInstrumentName As String
    Dim lsiListItem As ListItem
        
    Call UnselectAllInstruments
    
    sInstrumentName = moInstrumentManager.GetInstrumentName(vlInstrumentNumber)
    lvwInstruments.ListItems(sInstrumentName).Selected = True
    lvwInstruments.ListItems(sInstrumentName).EnsureVisible
    lvwInstruments.SetFocus
End Property


Private Sub lblAddRemoveFav_Click()
    Dim lsiSelectedListItem As ListItem
    Dim lListItemIndex As Long
    
    For lListItemIndex = lvwInstruments.ListItems.Count To 1 Step -1
        If lvwInstruments.ListItems(lListItemIndex).Selected Then
            If mbFavoritesMode Then
                Call RemoveFromFavorites(lvwInstruments.ListItems(lListItemIndex).SubItems(1))
            Else
                Call AddToFavorites(lvwInstruments.ListItems(lListItemIndex).SubItems(1))
            End If
        End If
    Next lListItemIndex
    Call FlashALable(lblAddRemoveFav)
End Sub

Private Sub lblFilter_Click()
    If mbListIsFiltred Then
        mbListIsFiltred = False
        Call LoadList
    Else
        Dim sWords As String
        sWords = InputBox("Enter words seperated by space which you want to be in intrument name:", "Enter criteria", vbNullString)
        Call FilterList(sWords)
        mbListIsFiltred = True
    End If
    Call ShowLables
End Sub

Private Sub lblOnOff_Click()
    Call lblPlayOnClick_Click
End Sub

Private Sub lblPlayOnClick_Click()
    mbPlayOnClick = Not mbPlayOnClick
    Call ShowLables
End Sub

Private Sub lblShowHideFav_Click()
    Call SetFavoriteInstrumentsMode(Not mbFavoritesMode)
End Sub

Private Sub lvwInstruments_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvwInstruments.SortKey = ColumnHeader.Position - 1
    If lvwInstruments.SortOrder = lvwAscending Then
        lvwInstruments.SortOrder = lvwDescending
    Else
        lvwInstruments.SortOrder = lvwAscending
    End If
    lvwInstruments.Sorted = True
End Sub

Private Sub lvwInstruments_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim lsiSelectedListItem As ListItem
    Set lsiSelectedListItem = lvwInstruments.SelectedItem
    If Not (lsiSelectedListItem Is Nothing) Then
        RaiseEvent InstrumentSelected(CLng(lsiSelectedListItem.Text), lsiSelectedListItem.SubItems(1))
        If mbPlayOnClick Then
            mbSamplePlayStopped = False
            RaiseEvent PlaySample
        End If
    End If
End Sub

Private Sub lvwInstruments_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If mbPlayOnClick And (Not mbSamplePlayStopped) Then
        mbSamplePlayStopped = True
        RaiseEvent StopSamplePlay
    End If
End Sub

Private Sub lvwInstruments_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If mbPlayOnClick Then
        RaiseEvent StopSamplePlay
    End If
End Sub

Private Sub lvwInstruments_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
    Dim lsiSelectedListItem As ListItem
    Set lsiSelectedListItem = lvwInstruments.SelectedItem
    If Not (lsiSelectedListItem Is Nothing) Then
        Data.SetData "INSTRUMENT:" & CStr(Replace(lsiSelectedListItem.SubItems(1), " ", "_")) & " ", vbCFText
    End If
    AllowedEffects = vbDropEffectCopy
End Sub

Private Sub moInstrumentManager_InstrumentListChanged()
    Call LoadListEntries
End Sub

Private Sub moInstrumentManager_InstrumentNameChange(ByVal vlInstrumentNumber As Long, ByVal vsOldInstrumentName As String, ByVal vsNewInstrumentName As String, ByVal vsINIFileName As String, Cancle As Boolean)

    If ((mbFavoritesMode) And (vsINIFileName = moInstrumentManager.GetINIFileNameWithAutoPathAdd(m_sINIFileForFavInstruments)) _
        Or (Not mbFavoritesMode) And (vsINIFileName = moInstrumentManager.GetINIFileNameWithAutoPathAdd(m_sINIFileForInstruments))) Then
        If vsNewInstrumentName = vbNullString Then
            Call lvwInstruments.ListItems.Remove(vsOldInstrumentName)
        Else
            Call lvwInstruments.ListItems.Remove(vsOldInstrumentName)
            Call AddInstrumentToList(vlInstrumentNumber, vsNewInstrumentName, True)
        End If
    End If
End Sub

Private Sub UserControl_Initialize()
    Set moInstrumentManager = New InstrumentManager
    mbFavoritesMode = False
    mbListIsFiltred = False
    mbPlayOnClick = True
    mbSamplePlayStopped = True
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_bAppRelativePath = PropBag.ReadProperty("AppRelativePath", True)
    moInstrumentManager.AutoAddAppPath = m_bAppRelativePath
    m_sINIFileForInstruments = PropBag.ReadProperty("INIFileForInstruments", gsDEFAULT_INSTRUMENTS_INI)
    m_sINIFileForFavInstruments = PropBag.ReadProperty("INIFileForFavInstruments", gsDEFAULT_FAV_INSTRUMENTS_INI)
    Call ShowLables
    Call LoadList
End Sub

Public Property Get InstrumentManager() As InstrumentManager
    Set InstrumentManager = moInstrumentManager
End Property

Private Sub UserControl_Resize()
    
    Dim lLablesWidth As Long

    lLablesWidth = linLablesBottom.Y1 - linLablesTop.Y1 + 1

    lvwInstruments.Left = 0
    lvwInstruments.Top = 0
    lvwInstruments.Height = Height - lLablesWidth
    lvwInstruments.Width = Width
    lvwInstruments.ColumnHeaders(1).Width = CLng(TextWidth("999999"))
    lvwInstruments.ColumnHeaders(2).Width = CLng(TextWidth("XXXXXXXXXXXXXXXXXXXXXXXXXXXX"))
    Dim lLastHeaderWidth As Long
    lLastHeaderWidth = Width - lvwInstruments.ColumnHeaders(1).Width - lvwInstruments.ColumnHeaders(2).Width - 100
    If lLastHeaderWidth > 0 Then
        lvwInstruments.ColumnHeaders(3).Width = lLastHeaderWidth
    Else
        lvwInstruments.ColumnHeaders(3).Width = 100
    End If
    
    linLablesTop.Y1 = lvwInstruments.Height + 10
    linLablesBottom.Y1 = linLablesTop.Y1 + lLablesWidth
    
    lblShowHideFav.Top = linLablesTop.Y1
    lblAddRemoveFav.Top = linLablesTop.Y1
    lblFilter.Top = linLablesTop.Y1
    lblPlayOnClick.Top = linLablesTop.Y1
    lblOnOff.Top = lblPlayOnClick.Top + lblPlayOnClick.Height - lblOnOff.Height + 10
    
    lblFilter.Left = Width - lblFilter.Width
    lblAddRemoveFav.Left = lblFilter.Left - lblAddRemoveFav.Width - 50
    lblShowHideFav.Left = lblAddRemoveFav.Left - lblShowHideFav.Width - 50
    lblPlayOnClick.Left = lblShowHideFav.Left - lblPlayOnClick.Width - 50
    lblOnOff.Left = lblPlayOnClick.Left + lblPlayOnClick.Width - lblOnOff.Width + 10
End Sub

Public Property Get INIFileForInstruments() As String
    INIFileForInstruments = m_sINIFileForInstruments
End Property

Public Property Let INIFileForInstruments(ByVal vsFileName As String)
    m_sINIFileForInstruments = vsFileName
    PropertyChanged ("INIFileForInstruments")
    moInstrumentManager.INIFileName = m_sINIFileForInstruments
End Property

Public Property Get INIFileForFavInstruments() As String
    INIFileForFavInstruments = m_sINIFileForFavInstruments
End Property

Public Property Let INIFileForFavInstruments(ByVal vsFileName As String)
    m_sINIFileForFavInstruments = vsFileName
    PropertyChanged ("INIFileForFavInstruments")
End Property

Public Property Get AppRelativePath() As Boolean
    AppRelativePath = m_bAppRelativePath
End Property

Public Property Let AppRelativePath(ByVal vboolAppRelativePath As Boolean)
    m_bAppRelativePath = vboolAppRelativePath
    moInstrumentManager.AutoAddAppPath = m_bAppRelativePath
    PropertyChanged ("INIFileForInstruments")
End Property
Private Sub UserControl_Terminate()
    Set moInstrumentManager = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "INIFileForInstruments", m_sINIFileForInstruments, gsDEFAULT_INSTRUMENTS_INI
    PropBag.WriteProperty "INIFileForFavInstruments", m_sINIFileForFavInstruments, gsDEFAULT_FAV_INSTRUMENTS_INI
    PropBag.WriteProperty "AppRelativePath", m_bAppRelativePath, True
End Sub

Private Sub LoadListEntries()
    
    Dim lInstrumentIndex As Long
    Dim lsiListItem As ListItem
    Dim sInstrumentName As String
    Dim lInstrumentNumber As Long
    
    lvwInstruments.ListItems.Clear
    
    For lInstrumentIndex = 1 To moInstrumentManager.InstrumentCount
            sInstrumentName = moInstrumentManager.GetInstrumentName(lInstrumentIndex, True)
            lInstrumentNumber = moInstrumentManager.GetInstrumentNumber(sInstrumentName)
            If Trim$(sInstrumentName) <> vbNullString Then
                Call AddInstrumentToList(lInstrumentNumber, sInstrumentName)
            End If
    Next lInstrumentIndex
    lvwInstruments.Sorted = False
    lvwInstruments.Sorted = True
End Sub

Public Sub SetFavoriteInstrumentsMode(ByVal vboolIsFavoritesInstrumentsMode As Boolean)
    mbFavoritesMode = vboolIsFavoritesInstrumentsMode
    Call ShowLables
    Call LoadList
End Sub

Public Sub AddToFavorites(ByVal vsInstrumentName As String)
    Dim lInstrumentNumber As Long
    lInstrumentNumber = moInstrumentManager.GetInstrumentNumber(vsInstrumentName)
    Call moInstrumentManager.SetInstrumentName(lInstrumentNumber, vsInstrumentName, m_sINIFileForFavInstruments)
End Sub

Public Sub RemoveFromFavorites(ByVal vsInstrumentName As String)
    Dim lInstrumentNumber As Long
    lInstrumentNumber = moInstrumentManager.GetInstrumentNumber(vsInstrumentName)
    Call moInstrumentManager.SetInstrumentName(lInstrumentNumber, vbNullString, m_sINIFileForFavInstruments)
End Sub

Public Sub FilterList(ByVal vsWords As String)
    Call moInstrumentManager.FilterInstrumentList(vsWords)
End Sub

Public Sub LoadList()
    If mbFavoritesMode Then
        moInstrumentManager.INIFileName = m_sINIFileForFavInstruments
    Else
        moInstrumentManager.INIFileName = m_sINIFileForInstruments
    End If
End Sub

Private Sub ShowLables()
    If mbFavoritesMode Then
        lblShowHideFav.Caption = "4"
        lblShowHideFav.ToolTipText = "Show all instruments"
        lblAddRemoveFav.Caption = Chr(251)
        lblAddRemoveFav.ToolTipText = "Remove from favorites)"
    Else
        lblShowHideFav.Caption = "2"
        lblShowHideFav.ToolTipText = "Show only favorites instruments"
        lblAddRemoveFav.Caption = Chr(252)
        lblAddRemoveFav.ToolTipText = "Add to favorites"
    End If
    
    If mbListIsFiltred Then
        lblFilter.Caption = "x"
        lblFilter.ToolTipText = "Remove Filter"
    Else
        lblFilter.Caption = "y"
        lblFilter.ToolTipText = "Filter the list"
    End If
    
    If Not mbPlayOnClick Then
        lblOnOff.Caption = Chr$(253)
        lblOnOff.ForeColor = vbRed
        lblOnOff.Font.Bold = False
        lblPlayOnClick.ToolTipText = "Don't play instrument while selecting"
        lblOnOff.ToolTipText = lblPlayOnClick.ToolTipText
    Else
        lblOnOff.Caption = Chr$(254)
        lblOnOff.ForeColor = RGB(0, 128, 0)
        lblOnOff.Font.Bold = True
        lblPlayOnClick.ToolTipText = "Play instrument while selecting"
        lblOnOff.ToolTipText = lblPlayOnClick.ToolTipText
    End If
    
End Sub

Private Sub FlashALable(ByVal vlblLable As Label)
    Dim lOriginalColor As Long
    lOriginalColor = vlblLable.ForeColor
    vlblLable.ForeColor = vbRed
    DoEvents
    Call Sleep(400)
    vlblLable.ForeColor = lOriginalColor
End Sub

Private Sub AddInstrumentToList(ByVal vlInstrumentNumber As Long, ByVal vsInstrumentName As String, Optional ByVal vboolSelectAddedInstrument As Boolean = False)
    Dim lsiListItem As ListItem
    Dim sInstrumentType As String
    Set lsiListItem = lvwInstruments.ListItems.Add(, vsInstrumentName, Format$(vlInstrumentNumber, "000"))
    lsiListItem.SubItems(1) = vsInstrumentName
    
    Select Case moInstrumentManager.GetInstrumentType(vlInstrumentNumber)
        Case itpBass
            sInstrumentType = "Bass"
        Case itpBrass
            sInstrumentType = "Brass"
        Case itpChromaticPercussion
            sInstrumentType = "Chromatic Percussion"
        Case itpEnsemble
            sInstrumentType = "Ensemble"
        Case itpGuitar
            sInstrumentType = "Guitar"
        Case itpOrgan
            sInstrumentType = "Organ"
        Case itpPiano
            sInstrumentType = "Organ"
        Case itpPipe
            sInstrumentType = "Pipe"
        Case itpReed
            sInstrumentType = "Reed"
        Case itpSoundEffects
            sInstrumentType = "Sound Effects"
        Case itpSynthLead
            sInstrumentType = "Synthesizer Lead"
        Case itpSynthPad
            sInstrumentType = "Synthesizer Pad"
        Case itpStrings
            sInstrumentType = "Strings"
        Case itpFX
            sInstrumentType = "FX"
        Case itpMisc1
            sInstrumentType = "Misc"
        Case itpMisc2
            sInstrumentType = "Misc"
        Case Else
            sInstrumentType = "Unrecognized(" & moInstrumentManager.GetInstrumentType(vlInstrumentNumber) & ")"
    End Select
    
    lsiListItem.SubItems(2) = sInstrumentType
    
    If vboolSelectAddedInstrument = True Then
        Call UnselectAllInstruments
        lsiListItem.Selected = True
        lsiListItem.EnsureVisible
    End If
End Sub

Private Sub UnselectAllInstruments()
    Dim lsiListItem As ListItem
    
    For Each lsiListItem In lvwInstruments.ListItems
        If lsiListItem.Selected Then
            lsiListItem.Selected = False
        End If
    Next lsiListItem
End Sub
