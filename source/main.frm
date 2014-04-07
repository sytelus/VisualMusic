VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visual Music"
   ClientHeight    =   8310
   ClientLeft      =   -210
   ClientTop       =   615
   ClientWidth     =   9465
   Icon            =   "main.frx":0000
   KeyPreview      =   -1  'True
   LinkMode        =   1  'Source
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   9465
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrFirstTime 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   4080
      Top             =   3960
   End
   Begin VB.Timer tmrInstructionProcesor 
      Left            =   180
      Top             =   1680
   End
   Begin VB.TextBox BaseCommands 
      Height          =   285
      Left            =   840
      TabIndex        =   11
      Text            =   "Text box for DDE link"
      Top             =   1680
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Timer tmrSampleNote 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   660
      Top             =   1620
   End
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   60
      TabIndex        =   7
      Top             =   1980
      Width           =   9315
      Begin VisualMusic.MIDIParamControls MIDIParamControls1 
         Height          =   2355
         Left            =   4740
         TabIndex        =   8
         Top             =   300
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   4154
      End
      Begin VisualMusic.InstrumentSelector InstrumentSelector1 
         Height          =   2670
         Left            =   60
         TabIndex        =   2
         Top             =   420
         Width           =   4200
         _ExtentX        =   7408
         _ExtentY        =   4710
         INIFileForInstruments=   ""
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&Instruments:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   180
         Width           =   855
      End
   End
   Begin MSComctlLib.Toolbar tlbPlay 
      Height          =   570
      Left            =   8580
      TabIndex        =   6
      Top             =   5460
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   1005
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Style           =   1
      ImageList       =   "ilsMain"
      HotImageList    =   "ilsMain"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btnPlay"
            Object.ToolTipText     =   "Play the currently selected tune"
            ImageKey        =   "keyPlay"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btnStop"
            Object.ToolTipText     =   "Stop Play/Recording"
            ImageKey        =   "keyStop"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btnRecord"
            Object.ToolTipText     =   "Record whatever you play"
            ImageKey        =   "keyRecord"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btnPlayAll"
            Object.ToolTipText     =   "Play all the tunes simultaneously"
            ImageKey        =   "keyPlayAll"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsMain 
      Left            =   8820
      Top             =   -240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":1CFA
            Key             =   "keyNew"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":2B4E
            Key             =   "keyOpen"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":39A2
            Key             =   "keySave"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":56AE
            Key             =   "keyHelp"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":6502
            Key             =   "keyPlay"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":6DDE
            Key             =   "keyStop"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":7C32
            Key             =   "keyRecord"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":993E
            Key             =   "keyPlayAll"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgFileSave 
      Left            =   8280
      Top             =   -180
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Save To MScript File"
      FileName        =   "*.MSC"
      Filter          =   "MScript File (*.MSC)|*.msc|All Files (*.*)|*.*"
      Flags           =   2054
   End
   Begin MSComDlg.CommonDialog dlgFileOpn 
      Left            =   7800
      Top             =   -180
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Open MScript File"
      FileName        =   "*.MSC"
      Filter          =   "MScript File (*.MSC)|*.msc|All Files (*.*)|*.*"
      Flags           =   4100
   End
   Begin VisualMusic.MScriptEditor MScriptEditor1 
      Height          =   2730
      Left            =   60
      TabIndex        =   4
      Top             =   5460
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   4815
   End
   Begin VisualMusic.MIDIKeyboard kbdMain 
      Height          =   975
      Left            =   60
      TabIndex        =   0
      Top             =   660
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   1720
      ShowFocusIndicator=   0   'False
   End
   Begin MSComctlLib.Toolbar tlbMain 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   1058
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "ilsMain"
      HotImageList    =   "ilsMain"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btnNew"
            Object.ToolTipText     =   "New composition"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btnOpen"
            Object.ToolTipText     =   "Open composition from file"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btnSave"
            Object.ToolTipText     =   "Save composition to file"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btnHelp"
            Object.ToolTipText     =   "Help!"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Label lblScriptHeading 
      AutoSize        =   -1  'True
      Caption         =   "Script:"
      Height          =   195
      Left            =   2280
      TabIndex        =   10
      Top             =   5220
      UseMnemonic     =   0   'False
      Width           =   450
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Tunes:"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   5220
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "C  D   E  F  G  A  B  C  D  E  F  G  A  B  C  D   E  F  G  A  B  C  D  E  F  G  A  B  C  "
      Height          =   195
      Left            =   2040
      TabIndex        =   3
      Top             =   1680
      Width           =   5805
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditCopyTunes 
         Caption         =   "&Copy Tunes"
      End
      Begin VB.Menu mnuEditPasteTunes 
         Caption         =   "&Paste Tunes"
      End
      Begin VB.Menu mnuEditBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOption 
         Caption         =   "&Option"
         Begin VB.Menu mnuVisualFeedBack 
            Caption         =   "&Visual Feedback For Playback"
         End
         Begin VB.Menu mnuUseProfessionalKeyboardLayout 
            Caption         =   "&Use Professional Keyboard Layout"
         End
         Begin VB.Menu mnuEnableMultiKeyRecording 
            Caption         =   "&Enable Multi-Key Recording"
         End
      End
   End
   Begin VB.Menu mnuTune 
      Caption         =   "&Tune"
      Begin VB.Menu mnuPlay 
         Caption         =   "&Play Selected Tune"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuRecord 
         Caption         =   "&Record"
      End
      Begin VB.Menu mnuPlayAll 
         Caption         =   "Play &All Tunes"
      End
      Begin VB.Menu mnuCompile 
         Caption         =   "&Compile Selected Tune"
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu mnuTuneBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStop 
         Caption         =   "S&top Play/Recording"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpMScriptReference 
         Caption         =   "&MScript Reference..."
      End
      Begin VB.Menu mnuHelpKeyboardShortcuts 
         Caption         =   "&Keyboard Shortcuts..."
      End
      Begin VB.Menu mnuHelpBreak1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuJoinMailingList 
         Caption         =   "&Join Mailing List"
      End
      Begin VB.Menu mnuSendYourTunes 
         Caption         =   "Send &Your Tunes"
      End
      Begin VB.Menu mnuHelpSendFeedback 
         Caption         =   "Send &Feedback"
      End
      Begin VB.Menu mnuHelpBreak2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpHomePage 
         Caption         =   "Visual Music &Homepage"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuHelpShital 
         Caption         =   "&Shital's Homepage"
      End
      Begin VB.Menu mnuHelpBreak3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Dim WithEvents oMidiEng As MIDIEngine
Attribute oMidiEng.VB_VarHelpID = -1
Dim WithEvents oMIDIEngForFeedback As MIDIEngine
Attribute oMIDIEngForFeedback.VB_VarHelpID = -1
Dim oClock As ClockProvider
Dim oMScriptCompiler As New MScriptCompiler
Dim WithEvents oMScriptRecorder As MScriptRecorder
Attribute oMScriptRecorder.VB_VarHelpID = -1
'Dim WithEvents oTimer As clsTimer
Private WithEvents moclMScripts As MScripts
Attribute moclMScripts.VB_VarHelpID = -1
Private moHTMLHelp As clsHTMLHelp

Private Const msPROG_NAME As String = "Visual Music"
Private Const mlSAMPLE_NOTE1 As Long = 13
Private Const mlSAMPLE_NOTE2 As Long = 15
Private Const mlSAMPLE_NOTE3 As Long = 17

Private Const sBUTTON_NEW As String = "btnNew"
Private Const sBUTTON_OPEN As String = "btnOpen"
Private Const sBUTTON_SAVE As String = "btnSave"
Private Const sBUTTON_HELP As String = "btnHelp"
Private Const sBUTTON_PLAY As String = "btnPlay"
Private Const sBUTTON_STOP As String = "btnStop"
Private Const sBUTTON_RECORD As String = "btnRecord"
Private Const sBUTTON_PLAY_ALL As String = "btnPlayAll"

Private Const sDDE_COMMAND_ACTIVATE As String = "ACTIVATE"
Private Const sDDE_COMMAND_EXIT As String = "EXIT"
Private Const sDDE_COMMAND_PLAY As String = "PLAY"
Private Const sDDE_COMMAND_PLAY_ALL As String = "PLAY_ALL"
Private Const sDDE_COMMAND_STOP_PLAY As String = "STOP_PLAY"
Private Const sDDE_COMMAND_PAUSE As String = "PAUSE"
Private Const sDDE_COMMAND_LOAD As String = "LOAD"
Private Const sDDE_COMMAND_SAVE As String = "SAVE"
Private Const sDDE_COMMAND_SAVE_AS As String = "SAVE_AS"
Private Const sDDE_COMMAND_NEW As String = "NEW"
Private Const sDDE_COMMAND_SHOW_HELP As String = "SHOW_HELP"

Private Const msDEFAULT_TUNE_NAME As String = "<New Tune>"
Private Const msREG_APP_ROOT As String = "Visual Music"
Private Const msREG_SECTION_OPTION As String = "Options"
Private Const msREG_SECTION_STATE As String = "State"
Private Const msREG_KEY_EXT_REGISTERED As String = "ExtAndIconRegistered"
Private Const msREG_KEY_VISUAL_FEEDBACK As String = "VisualFeedBack"
Private Const msREG_KEY_IS_FIRST_TIME As String = "IsFirstTime"
Private Const msREG_KEY_USE_PROFESSIONAL_KEYBOARD_LAYOUT As String = "UseProfessionalKeyboardLayout"
Private Const msREG_KEY_ENABLE_MULTI_KEY_RECORDING As String = "EnableMultiKeyRecording"

Private mlLastSelectedItemInList As Long
Private mbVisualFeedBack As Boolean
Private mbProfessionalKeyboardLayout As Boolean
Private mbEnableMultiKeyRecording As Boolean

Private m_sFileName As String
Private mbFileSaved As Boolean
Private mbIsRunningFirstTime As Boolean

Private bRecordingWasPausedWhilePlayingSample As Boolean

Const mlNOTE As Long = 13

Private Sub btnCompileAll_Click()
    MScriptEditor1.MScripts.CompileAll
End Sub

Private Sub btnPlayAll_Click()
    Call MScriptEditor1.MScripts.PlayAll
End Sub

Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    oMidiEng.StopNote mlNOTE
End Sub

Private Sub chkKeybordTurnOff_Click()

End Sub

Private Sub Form_Activate()

    On Error GoTo ERR_Form_Activate

    Static bDoneOnce As Boolean
    
    If Not bDoneOnce Then
        Call MScriptEditor1.SelectTune(1, True)
        If InstrumentSelector1.InstrumentManager.IsInstrumentExist(oMidiEng.Instrument) Then
            InstrumentSelector1.SelectedInstrument = oMidiEng.Instrument
        End If
        'Set oMScriptCompiler.InstrumentManager = InstrumentSelector1.InstrumentManager
        
        'Reflact defaults
        Set oMIDIEngForFeedback = oMidiEng
        oMIDIEngForFeedback.RaisePropertyChangeEvents
        Set oMIDIEngForFeedback = Nothing
        
        bDoneOnce = True
        If mbIsRunningFirstTime Then
            Call FillSampleScript
            Call SaveSetting(msREG_APP_ROOT, msREG_SECTION_STATE, msREG_KEY_IS_FIRST_TIME, "False")
            tmrFirstTime.Enabled = True
        End If
    End If

Exit Sub
ERR_Form_Activate:
    ShowError
End Sub

Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)

    On Error GoTo ERR_Form_LinkExecute

    Cancel = False
    Me.SetFocus
    
    Select Case UCase$(CmdStr)
        Case sDDE_COMMAND_ACTIVATE
            Me.SetFocus
        Case sDDE_COMMAND_EXIT
            Unload Me
            End
        Case sDDE_COMMAND_NEW
            Call mnuFileNew_Click
        Case sDDE_COMMAND_PAUSE
            err.Raise 1000, , "DDE command " & CmdStr & " Not yet implemented"
        Case sDDE_COMMAND_PLAY_ALL
            Call mnuPlayAll_Click
        Case sDDE_COMMAND_SAVE
            Call mnuFileSave_Click
        Case sDDE_COMMAND_STOP_PLAY
            Call mnuStop_Click
        Case Else
            'Find the "("
            Dim lFirstBracketPos As Long
            Dim lLastBracketPos As Long
            Dim sParam As String
            Dim sCommand As String
            lFirstBracketPos = InStr(1, CmdStr, "(")
            If lFirstBracketPos <> 0 Then
                lLastBracketPos = InStrRev(CmdStr, ")")
                If lLastBracketPos <> 0 Then
                    If lFirstBracketPos = lLastBracketPos - 1 Then
                        sParam = vbNullString
                    Else
                        sParam = Mid$(CmdStr, lFirstBracketPos + 1, lLastBracketPos - lFirstBracketPos - 1)
                    End If
                    sCommand = Trim$(Mid$(CmdStr, 1, lFirstBracketPos - 1))
                    Select Case UCase$(sCommand)
                    Case sDDE_COMMAND_LOAD
                        If sParam = vbNullString Then
                            Call mnuFileOpen_Click
                        Else
                            FileName = sParam
                            Call LoadMScriptsFromFile(sParam)
                        End If
                    Case sDDE_COMMAND_SAVE_AS
                            FileName = sParam
                            Call SaveMScriptsToFile(sParam)
                    Case sDDE_COMMAND_SHOW_HELP
                            If sParam <> vbNullString Then
                                Call moHTMLHelp.ShowTopic(sParam)
                            Else
                                moHTMLHelp.ShowContent
                            End If
                    Case sDDE_COMMAND_PLAY
                        MScriptEditor1.SaveCurrentText
                        If sParam = vbNullString Then
                            MScriptEditor1.PlaySelectedTune
                        Else
                            Call moclMScripts(sParam).Play
                        End If
                    End Select
                Else
                    Cancel = True
                    err.Raise 1000, , "Missing Parathenesis ')' in " & CmdStr
                End If
            Else
                Cancel = True
                err.Raise 1000, , "Unrecognised DDE command received: " & CmdStr
            End If
    End Select
    
Exit Sub
ERR_Form_LinkExecute:
    ShowError
End Sub

Private Sub Form_Load()
    
    On Error GoTo ERR_Form_Load
    
    Const sIMAGE_KEY_FOR_NEW As String = "keyNew"
    Const sIMAGE_KEY_FOR_OPEN As String = "keyOpen"
    Const sIMAGE_KEY_FOR_SAVE As String = "keySave"
    Const sIMAGE_KEY_FOR_HELP As String = "keyHelp"

    Dim bIsFileExtensionAndIconRegistered As Boolean
    bRecordingWasPausedWhilePlayingSample = False

    Call DDEHanlder

    Call CommandSwitchProcessor(Command$)
    
    mbIsRunningFirstTime = GetSetting(msREG_APP_ROOT, msREG_SECTION_STATE, msREG_KEY_IS_FIRST_TIME, "True")
    bIsFileExtensionAndIconRegistered = GetSetting(msREG_APP_ROOT, msREG_SECTION_STATE, msREG_KEY_EXT_REGISTERED, "False")
    mbVisualFeedBack = GetSetting(msREG_APP_ROOT, msREG_SECTION_OPTION, msREG_KEY_VISUAL_FEEDBACK, "True")
    mnuVisualFeedBack.Checked = mbVisualFeedBack
    mbProfessionalKeyboardLayout = GetSetting(msREG_APP_ROOT, msREG_SECTION_OPTION, msREG_KEY_USE_PROFESSIONAL_KEYBOARD_LAYOUT, "False")
    mnuUseProfessionalKeyboardLayout.Checked = mbProfessionalKeyboardLayout
    mbEnableMultiKeyRecording = GetSetting(msREG_APP_ROOT, msREG_SECTION_OPTION, msREG_KEY_ENABLE_MULTI_KEY_RECORDING, "True")
    mnuEnableMultiKeyRecording.Checked = mbEnableMultiKeyRecording
    
    If Not bIsFileExtensionAndIconRegistered Then
        Call RegisterFileExtensionAndIcon("MSC", "MScript File", GetPathWithSlash(App.Path) & App.EXEName & ".exe", GetPathWithSlash(App.Path) & "VisualMusic.ico")
        Call SaveSetting(msREG_APP_ROOT, msREG_SECTION_STATE, msREG_KEY_EXT_REGISTERED, "True")
    End If
    
    tlbMain.Buttons(sBUTTON_NEW).Image = sIMAGE_KEY_FOR_NEW
    tlbMain.Buttons(sBUTTON_OPEN).Image = sIMAGE_KEY_FOR_OPEN
    tlbMain.Buttons(sBUTTON_SAVE).Image = sIMAGE_KEY_FOR_SAVE
    tlbMain.Buttons(sBUTTON_HELP).Image = sIMAGE_KEY_FOR_HELP

    Set oMidiEng = New MIDIEngine
    
    Set moclMScripts = New MScripts
    oMidiEng.OpenOutputPort
    With oMidiEng
        .Instrument = 10
        .Octave = lMIDI_ENG_DEFAULT_OCTAVE
    End With
    
    With MIDIParamControls1
        .Octave = oMidiEng.Octave
        .Pan = oMidiEng.Pan
        .Volume = oMidiEng.Volume
    End With
    
'    Set oTimer = New clsTimer
'    With oTimer
'        .Interval = moclMScripts.TimerInterval
'        .hWnd = Me.hWnd
'        .Enabled = True
'    End With
    With tmrInstructionProcesor
        .Interval = moclMScripts.TimerInterval
    End With
    
    Set oClock = New ClockProvider
    
    Set moHTMLHelp = New clsHTMLHelp
    
    Dim i As Integer
    Dim sListItemText As String
    Set moclMScripts.InstructionClock = oClock
    Set moclMScripts.MScriptCompiler = oMScriptCompiler
    
    Set MScriptEditor1.MScripts = moclMScripts
    
    Set oMScriptRecorder = New MScriptRecorder
    Set oMScriptRecorder.ClockProvider = oClock
    Set oMScriptRecorder.MIDIEngine = oMidiEng
    oMScriptRecorder.EnableMultiKeyRecording = mbEnableMultiKeyRecording
    
    For i = 1 To 3
        sListItemText = "Tune" & i
        Call moclMScripts.Add("", sListItemText)
    Next i
    
    'Set keyboard layout
    Call SetKeyBoardLayout(mbProfessionalKeyboardLayout)
    
    InstrumentSelector1.INIFileForInstruments = "instruments.ini"
    
    If Command$ <> vbNullString Then
        FileName = Command$
        Call LoadMScriptsFromFile(FileName)
    End If
        
    Call RefreshCaption
    
    mbFileSaved = True
    
Exit Sub
ERR_Form_Load:
    Dim bIsMIDIError As Boolean
    bIsMIDIError = err.Number >= lERR_BASE_MIDIENG And err.Number <= lERR_MIDIENG_HIGHEST
    ShowError
    If bIsMIDIError Then
        Unload Me
        End
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Set oTimer = Nothing
    Set oMidiEng = Nothing
    Set oClock = Nothing
    Set moHTMLHelp = Nothing
End Sub

Private Sub HiTime1_Timer()
    On Error GoTo ERR_HiTime1_Timer
    oClock.GenerateClockPulse
Exit Sub
ERR_HiTime1_Timer:
    Call ShowError
End Sub

Private Sub InstrumentSelector1_InstrumentSelected(ByVal vlInstrumentNumber As Long, ByVal vsInstrumentName As String)
    oMidiEng.Instrument = vlInstrumentNumber
End Sub

Private Sub InstrumentSelector1_PlaySample()
    Call PauseRecordingBeforePlayingSample
    Call oMidiEng.PlayNote(mlSAMPLE_NOTE1)
    tmrSampleNote.Enabled = True
    tmrSampleNote.Tag = 1
End Sub

Private Sub InstrumentSelector1_StopSamplePlay()
    Call PauseRecordingBeforePlayingSample
    Call oMidiEng.StopNote(mlSAMPLE_NOTE1)
    Call oMidiEng.StopNote(mlSAMPLE_NOTE2)
    Call oMidiEng.StopNote(mlSAMPLE_NOTE3)
    Call StartRecordingAfterPlayingSample
End Sub

Private Sub kbdMain_KeyDown(ByVal vlKeyIndex As Long)
        oMidiEng.PlayNote vlKeyIndex
End Sub

Private Sub kbdMain_KeyUp(ByVal vlKeyIndex As Long)
    oMidiEng.StopNote vlKeyIndex
End Sub

Private Sub lstInstruments_GotFocus()
    'Panel3D3.SetFocus
End Sub

Private Sub MIDIParamControls1_SetOctave(rlOctave As Long)
    oMidiEng.Octave = rlOctave
End Sub

Private Sub MIDIParamControls1_SetPan(rlPan As Long)
    oMidiEng.Pan = rlPan
End Sub

Private Sub MIDIParamControls1_SetVolume(rlVolume As Long)
    oMidiEng.Volume = rlVolume
End Sub

Private Sub mnuCompile_Click()

    On Error GoTo ERR_mnuCompile_Click

    Dim sSelectedTune As String
    sSelectedTune = MScriptEditor1.CompileSelectedTune
    If sSelectedTune = vbNullString Then
        err.Raise 1000, , "No tune selected. First select the tune you want to compile"
    Else
        MsgBox sSelectedTune & " compiled successfully. Press Play button to run the tune."
    End If

Exit Sub
ERR_mnuCompile_Click:
    ShowError
End Sub

Private Sub mnuEdit_Click()
    Dim sClipboardText As String
    sClipboardText = Clipboard.GetText(vbCFText)
    If sClipboardText = vbNullString Then
        mnuEditPasteTunes.Enabled = False
    Else
        mnuEditPasteTunes.Enabled = True
    End If
End Sub

Private Sub mnuEditCopyTunes_Click()
    Dim sScriptText As String
    MScriptEditor1.SaveCurrentText
    sScriptText = moclMScripts.SaveToText
    Call Clipboard.Clear
    Call Clipboard.SetText(sScriptText)
End Sub

Private Sub mnuEditPasteTunes_Click()
    On Error GoTo ERR_mnuEditPasteTunes_Click
        
    Dim sClipboardText As String
    Dim oMScripts As MScripts
    
    Set oMScripts = New MScripts
    
    sClipboardText = Clipboard.GetText(vbCFText)
    
    Call oMScripts.LoadFromText(sClipboardText)
    
    Set oMScripts = Nothing
    
    'If no error occured
    Call moclMScripts.LoadFromText(sClipboardText)
    
Exit Sub
ERR_mnuEditPasteTunes_Click:
    Set oMScripts = Nothing
    Call ShowError
End Sub

Private Sub mnuEnableMultiKeyRecording_Click()
    mnuEnableMultiKeyRecording.Checked = Not mnuEnableMultiKeyRecording.Checked
    mbEnableMultiKeyRecording = mnuEnableMultiKeyRecording.Checked
    Call SaveSetting(msREG_APP_ROOT, msREG_SECTION_OPTION, msREG_KEY_ENABLE_MULTI_KEY_RECORDING, CStr(mbEnableMultiKeyRecording))
    If Not (oMScriptRecorder Is Nothing) Then
        oMScriptRecorder.EnableMultiKeyRecording = mbEnableMultiKeyRecording
    End If
End Sub

Private Sub mnuFileExit_Click()
    On Error GoTo ERR_mnuFileSaveAs_Click
    Call AskAndSaveIfDirty
    Unload Me
    End
Exit Sub
ERR_mnuFileSaveAs_Click:
    If err.Number <> cdlCancel Then
        Call ShowError
    End If
    End
End Sub

Private Sub mnuFileNew_Click()
    On Error GoTo ERR_mnuFileNew_Click
    Call FileNew
Exit Sub
ERR_mnuFileNew_Click:
    Call ShowError
End Sub

Private Sub mnuFileOpen_Click()
    On Error GoTo ERR_mnuFileOpen_Click
    Call FileOpen
Exit Sub
ERR_mnuFileOpen_Click:
    Call ShowError
End Sub

Private Sub mnuFileSave_Click()
    On Error GoTo ERR_mnuFileSave_Click
    Call FileSave
Exit Sub
ERR_mnuFileSave_Click:
    Call ShowError
End Sub

Private Sub mnuFileSaveAs_Click()
    On Error GoTo ERR_mnuFileSaveAs_Click
    Call FileSave(True)
Exit Sub
ERR_mnuFileSaveAs_Click:
    Call ShowError
End Sub

Private Sub mnuHelpAbout_Click()
    Call frmAbout.Show(vbModal, Me)
End Sub

Private Sub mnuHelpContents_Click()
    Call moHTMLHelp.ShowContent
End Sub

Private Sub mnuHelpHomePage_Click()
    Call moHTMLHelp.GotoWebpage("http://i.am/vmusic")
End Sub

Private Sub mnuHelpKeyboardShortcuts_Click()
    Call moHTMLHelp.ShowTopic("keyboardshortcuts.htm")
End Sub

Private Sub mnuHelpMScriptReference_Click()
    Call moHTMLHelp.ShowTopic("mscriptreference.htm")
End Sub

Private Sub mnuHelpSearch_Click()
    Call moHTMLHelp.ShowSearch
End Sub

Private Sub mnuHelpSendFeedback_Click()
    Call moHTMLHelp.InvokeEmailClient("shital_s@usa.net")
End Sub

Private Sub mnuHelpShital_Click()
    Call moHTMLHelp.GotoWebpage("http://i.am/shital")
End Sub

Private Sub mnuJoinMailingList_Click()
    Call moHTMLHelp.InvokeEmailClient("visualmusic-subscribe@egroups.com")
    MsgBox "Simply send a blank email to visualmusic-subscribe@egroups.com"
End Sub

Private Sub mnuPlay_Click()
    On Error GoTo ERR_mnuPlay_Click
    
    If Not oMScriptRecorder.IsRecording Then
        Dim sMScriptSelected As String
        sMScriptSelected = MScriptEditor1.GetSelectedTune
        If sMScriptSelected <> vbNullString Then
            Call MScriptEditor1.PlaySelectedTune
        Else
            err.Raise 1000, , "No tune selected for playing."
        End If
        Call MScriptEditor1.SetFocusToControl(mecTuneEditor)
    Else
        err.Raise 1000, , "First press Stop button to stop the recording"
    End If
Exit Sub
ERR_mnuPlay_Click:
    ShowError
End Sub

Private Sub mnuPlayAll_Click()

    On Error GoTo ERR_mnuPlayAll_Click

    MScriptEditor1.SaveCurrentText
    moclMScripts.PlayAll
    
Exit Sub
ERR_mnuPlayAll_Click:
    ShowError
End Sub

Private Sub mnuRecord_Click()
    
    On Error GoTo ERR_mnuRecord_Click

    If oMScriptRecorder.IsRecording Then
        oMScriptRecorder.StopRecording
    Else
        MScriptEditor1.Text = MScriptEditor1.Text & oMScriptRecorder.StartRecording()
    End If
    Call MScriptEditor1.SaveCurrentText
    Call kbdMain.SetFocus
        
Exit Sub
ERR_mnuRecord_Click:
    ShowError
End Sub

Private Sub mnuSendYourTunes_Click()
    Call moHTMLHelp.InvokeEmailClient("visualmusic@egroups.com")
    Call mnuEditCopyTunes_Click
    MsgBox "The script for your tunes is copied to clipboard." & vbCrLf & "Paste this tunes in your email client."
End Sub

Private Sub mnuStop_Click()

    On Error GoTo ERR_mnuStop_Click

    If oMScriptRecorder.IsRecording Then
        oMScriptRecorder.StopRecording
        Call MScriptEditor1.SaveCurrentText
    End If

    Dim oMScript As MScript
    For Each oMScript In moclMScripts
        oMScript.StopPlay
    Next oMScript
    
    Call oMidiEng.StopAllNotes
    
    Call kbdMain.MakeAllKeyDown
    
Exit Sub
ERR_mnuStop_Click:
    
End Sub

Private Sub mnuUseProfessionalKeyboardLayout_Click()
    mnuUseProfessionalKeyboardLayout.Checked = Not mnuUseProfessionalKeyboardLayout.Checked
    mbProfessionalKeyboardLayout = mnuUseProfessionalKeyboardLayout.Checked
    Call SaveSetting(msREG_APP_ROOT, msREG_SECTION_OPTION, msREG_KEY_USE_PROFESSIONAL_KEYBOARD_LAYOUT, CStr(mbProfessionalKeyboardLayout))
    Call SetKeyBoardLayout(mbProfessionalKeyboardLayout)
End Sub

Private Sub SetKeyBoardLayout(ByVal vboolProfessionalKeyboardLayout As Boolean)
    
    Call kbdMain.ClearKeyboardLayout
    
    If Not vboolProfessionalKeyboardLayout Then
        Call kbdMain.SetKeyCodeForIndex(vbKeyQ, 0)
        Call kbdMain.SetKeyCodeForIndex(vbKeyW, 1)
        Call kbdMain.SetKeyCodeForIndex(vbKeyE, 2)
        Call kbdMain.SetKeyCodeForIndex(vbKeyR, 3)
        Call kbdMain.SetKeyCodeForIndex(vbKeyT, 4)
        Call kbdMain.SetKeyCodeForIndex(vbKeyY, 5)
        Call kbdMain.SetKeyCodeForIndex(vbKeyU, 6)
        Call kbdMain.SetKeyCodeForIndex(vbKeyI, 7)
        Call kbdMain.SetKeyCodeForIndex(vbKeyO, 8)
        Call kbdMain.SetKeyCodeForIndex(vbKeyP, 9)
        Call kbdMain.SetKeyCodeForIndex(219, 10)
        Call kbdMain.SetKeyCodeForIndex(221, 11)
    Else
        Call kbdMain.SetKeyCodeForIndex(vbKeyA, 0)
        Call kbdMain.SetKeyCodeForIndex(vbKeyW, 1)
        Call kbdMain.SetKeyCodeForIndex(vbKeyS, 2)
        Call kbdMain.SetKeyCodeForIndex(vbKeyE, 3)
        Call kbdMain.SetKeyCodeForIndex(vbKeyD, 4)
        Call kbdMain.SetKeyCodeForIndex(vbKeyF, 5)
        Call kbdMain.SetKeyCodeForIndex(vbKeyT, 6)
        Call kbdMain.SetKeyCodeForIndex(vbKeyG, 7)
        Call kbdMain.SetKeyCodeForIndex(vbKeyY, 8)
        Call kbdMain.SetKeyCodeForIndex(vbKeyH, 9)
        Call kbdMain.SetKeyCodeForIndex(vbKeyU, 10)
        Call kbdMain.SetKeyCodeForIndex(vbKeyJ, 11)
        Call kbdMain.SetKeyCodeForIndex(vbKeyK, 12)
        Call kbdMain.SetKeyCodeForIndex(vbKeyO, 13)
        Call kbdMain.SetKeyCodeForIndex(vbKeyL, 14)
        Call kbdMain.SetKeyCodeForIndex(vbKeyP, 15)
        Call kbdMain.SetKeyCodeForIndex(186, 16)
        Call kbdMain.SetKeyCodeForIndex(192, 17)
        Call kbdMain.SetKeyCodeForIndex(221, 18)
    End If
End Sub

Private Sub mnuVisualFeedBack_Click()

    On Error GoTo ERR_mnuVisualFeedBack_Click

    mnuVisualFeedBack.Checked = Not mnuVisualFeedBack.Checked
    mbVisualFeedBack = mnuVisualFeedBack.Checked
    If Not mbVisualFeedBack Then
        Set oMIDIEngForFeedback = Nothing
    Else
        Call SetupMIDIFeedBack(MScriptEditor1.GetSelectedTune)
    End If
    Call SaveSetting(msREG_APP_ROOT, msREG_SECTION_OPTION, msREG_KEY_VISUAL_FEEDBACK, CStr(mbVisualFeedBack))

Exit Sub
ERR_mnuVisualFeedBack_Click:
    ShowError
End Sub

Private Sub moclMScripts_AddMScript(ByVal vsMScriptName As String, ByVal vsMScriptText As String)
    mbFileSaved = False
End Sub

Private Sub moclMScripts_DeleteAllMScript(rbCancel As Boolean)
    mbFileSaved = False
End Sub

Private Sub moclMScripts_DeleteMScript(ByVal vsMScriptName As String, rbCancel As Boolean)
    mbFileSaved = False
End Sub

Private Sub moclMScripts_EndRenameMScript(ByVal vsOldMScriptName As String, ByVal vsNewMScriptName As String)
    mbFileSaved = False
End Sub

Private Sub moclMScripts_InstructionProcessorError(ByVal vsMScriptName As String, ByVal Number As Long, ByVal Source As String, ByVal Description As String, ByVal HelpFile As String, ByVal HelpContext As Long)
    On Error GoTo ERR_moclMScripts_InstructionProcessorError
    
    err.Raise Number, Source, Description & " (" & "Tune:" & vsMScriptName & ")", HelpFile, HelpContext
    
Exit Sub
ERR_moclMScripts_InstructionProcessorError:
    ShowError
End Sub

Private Sub moclMScripts_InstructionProcessorExecutionStartStopEvent(ByVal vsMScriptName As String, ByVal vboolStartStopFlag As Boolean)
    If mbVisualFeedBack Then
        If MScriptEditor1.GetSelectedTune = vsMScriptName Then
            Call SetupMIDIFeedBack(vsMScriptName)
        End If
    Else
        Set oMIDIEngForFeedback = Nothing
    End If
End Sub

Private Sub moclMScripts_ModifyMScriptText(ByVal vsMScriptName As String, ByVal vsOldMScriptText As String, ByVal vsNewMScriptText As String, rbCancel As Boolean)
    mbFileSaved = False
End Sub

Private Sub moclMScripts_PrintRequest(ByVal vsMScriptName As String, ByVal vsStringToPrint As String)
    frmDebug.txtDebug.Text = frmDebug.txtDebug.Text & vsMScriptName & " : " & vsStringToPrint & vbCrLf
    If Not frmDebug.Visible Then
        frmDebug.Show , Me
    End If
End Sub

Private Sub MScriptEditor1_TuneSelected(ByVal vsTuneName As String)
    If vsTuneName <> vbNullString Then
        lblScriptHeading.Caption = "Script for " & vsTuneName & ":"
        tlbPlay.Buttons(sBUTTON_PLAY).ToolTipText = "Play " & vsTuneName
        Call SetupMIDIFeedBack(vsTuneName)
    Else
        lblScriptHeading.Caption = "Script:"
        tlbPlay.Buttons(sBUTTON_PLAY).ToolTipText = "Play the currently selected tune"
    End If
End Sub

Private Sub SetupMIDIFeedBack(ByVal vsTuneName As String)
        If mbVisualFeedBack And (vsTuneName <> vbNullString) Then
            If Not (moclMScripts(vsTuneName).InstructionProcessor Is Nothing) Then
                If moclMScripts(vsTuneName).InstructionProcessor.IsExecuting Then
                    Set oMIDIEngForFeedback = moclMScripts(vsTuneName).InstructionProcessor.moMIDIEngine
                    Call kbdMain.MakeAllKeyDown
                    oMIDIEngForFeedback.RaisePropertyChangeEvents
                Else
                    Set oMIDIEngForFeedback = Nothing
                End If
            Else
                Set oMIDIEngForFeedback = Nothing
            End If
        Else
            Set oMIDIEngForFeedback = Nothing
        End If
End Sub

Private Sub oMIDIEngForFeedback_NoteStarted(ByVal vlNoteNumber As Long, ByVal vlOctave As Long, ByVal vlVolume As Long)
    Call kbdMain.MakeKeyDown(vlNoteNumber, False, True)
End Sub

Private Sub oMIDIEngForFeedback_NoteStoped(ByVal vlNoteNumber As Long, ByVal vlOctave As Long)
    Call kbdMain.MakeKeyUp(vlNoteNumber, False, True)
End Sub

Private Sub oMIDIEngForFeedback_PropertyChanged(ByVal venmProperty As MIDIEngineProperties, ByVal rlOldValue As Long, rlNewValue As Long, rbIgnoreNewValue As Boolean)
    Select Case venmProperty
        Case mepInstrument
            If InstrumentSelector1.INIFileForInstruments <> vbNullString Then
                InstrumentSelector1.SelectedInstrument = rlNewValue
            End If
        Case mepOctave
            MIDIParamControls1.Octave = rlNewValue
        Case mepPan
            MIDIParamControls1.Pan = rlNewValue
        Case mepVolume
            MIDIParamControls1.Volume = rlNewValue
    End Select
End Sub

Private Sub oMScriptRecorder_MScriptTextChanged(ByVal vboolIsAppended As Boolean, ByVal vsAppendedText As String)
    If vboolIsAppended Then
        MScriptEditor1.Text = MScriptEditor1.Text & vsAppendedText
    Else
        MScriptEditor1.Text = oMScriptRecorder.MScriptText
    End If
End Sub

'Private Sub oTimer_Timer()
'    On Error GoTo ERR_oTimer_Timer
'    oClock.GenerateClockPulse
'Exit Sub
'ERR_oTimer_Timer:
'    Call ShowError
'End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo ERR_Form_KeyDown

    If Not (KeyCode = 0 And Shift = 0) Then

        Select Case KeyCode
            Case vbKeyF5 And (Shift = 0)
'                Call MScriptEditor1.PlaySelectedTune
'                KeyCode = 0
            Case vbKeyF5 And ((Shift And vbCtrlMask) <> 0)
                Call mnuCompile_Click
                KeyCode = 0
            Case vbKeyS And ((Shift And vbCtrlMask) <> 0)
                Call FileSave
                KeyCode = 0
            Case vbKeyI And ((Shift And vbAltMask) <> 0)
                Call InstrumentSelector1.SetFocus
                KeyCode = 0
            Case vbKeyT And ((Shift And vbAltMask) <> 0)
                MScriptEditor1.SetFocus
                Call MScriptEditor1.SetFocusToControl(mecTuneList)
                KeyCode = 0
            Case vbKeyS And ((Shift And vbAltMask) <> 0)
                MScriptEditor1.SetFocus
                Call MScriptEditor1.SetFocusToControl(mecTuneEditor)
                KeyCode = 0
            Case Else
                If IsKeyForMIDIKeyboard(KeyCode, Shift) Then
                    Call kbdMain.MakeKeyDown(KeyCode, True)
                    KeyCode = 0
                Else    'If (Not (TypeOf ActiveControl Is MIDIParamControls)) Then
                    If MIDIParamControls1.ProcessKey(KeyCode, Shift) = True Then
                        KeyCode = 0
                    ElseIf MScriptEditor1.ProcessOnKeyDown(KeyCode, Shift) Then
                        KeyCode = 0
                    End If
                End If
        End Select
    End If
    
Exit Sub
ERR_Form_KeyDown:
    Call ShowError
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If IsKeyForMIDIKeyboard(KeyCode, Shift) Then
        Call kbdMain.MakeKeyUp(KeyCode, True)
        KeyCode = 0
    End If
End Sub

Private Sub ShowError(Optional ByVal vsMScriptName As String = vbNullString)
    Call DisplayErrorMessage
    Dim sErrorSource As String
    sErrorSource = err.Source
    Call MoveCursorToErrorPos(sErrorSource, True, vsMScriptName)
End Sub

Private Sub MoveCursorToErrorPos(ByVal vsErrorSource As String, Optional ByVal vboolIgnoreErrors As Boolean = False, Optional ByVal vsMScriptName As String = vbNullString)
    
    On Error GoTo ERR_MoveCursorToErrorPos
    
    Dim lColonPos As Long
    Dim lErrorLineNumber As Long
    
    lColonPos = InStr(1, vsErrorSource, ":")
    If lColonPos <> 0 Then
        lErrorLineNumber = Mid$(vsErrorSource, lColonPos + 1)
        If vsMScriptName <> vbNullString Then
            Call MScriptEditor1.SelectTune(vsMScriptName, True)
        Else
            Call MScriptEditor1.SelectTune(moclMScripts.LastCompiledMScript, True)
        End If
        Call MScriptEditor1.SelectWordAt(lErrorLineNumber)
    End If

Exit Sub
ERR_MoveCursorToErrorPos:
    If Not vboolIgnoreErrors Then
        Call ReRaiseError
    End If
End Sub

Private Sub RefreshCaption()
    Me.Caption = msPROG_NAME & " " & "Ver." & App.Major & "." & App.Minor & "." & App.Revision & " - " & AlternateStrIfNull(FileName, msDEFAULT_TUNE_NAME)
End Sub

Private Property Get FileName() As String
    FileName = m_sFileName
End Property

Private Property Let FileName(vsFileName As String)
    m_sFileName = vsFileName
    Call RefreshCaption
End Property

Private Sub FileNew()
    Call MScriptEditor1.SaveCurrentText
    Call AskAndSaveIfDirty
    Call moclMScripts.Clear
    mbFileSaved = True
    FileName = vbNullString
Exit Sub
ERR_FileNew:
    If err.Number <> cdlCancel Then
        ReRaiseError
    End If
End Sub

Private Sub FileOpen()
    
    On Error GoTo ERR_FileOpen
    
    Static bOpenDialogCalledOnce As Boolean
    
    Call MScriptEditor1.SaveCurrentText
    Call AskAndSaveIfDirty
    
    If Not bOpenDialogCalledOnce Then
        dlgFileOpn.InitDir = App.Path
    Else
        dlgFileOpn.InitDir = vbNullString
    End If
    dlgFileOpn.ShowOpen
    FileName = dlgFileOpn.FileName
    Call LoadMScriptsFromFile(FileName)
    Call MScriptEditor1.SelectTune(1, True)
    If (Not bOpenDialogCalledOnce) And mbIsRunningFirstTime Then
        MsgBox "Press Play button on the right to play this file.", , "How To..."
    End If
    bOpenDialogCalledOnce = True
Exit Sub
ERR_FileOpen:
    If err.Number <> cdlCancel Then
        ReRaiseError
    End If
End Sub

Private Sub FileSave(Optional ByVal vboolShowSaveDlg As Boolean = False)
    
    On Error GoTo ERR_FileSave
    
    If (FileName = vbNullString) Or (vboolShowSaveDlg = True) Then
        dlgFileSave.ShowSave
        FileName = dlgFileSave.FileName
    End If
    Call SaveMScriptsToFile(FileName)

Exit Sub
ERR_FileSave:
    If err.Number <> cdlCancel Then
        ReRaiseError
    End If
End Sub

Private Sub LoadMScriptsFromFile(ByVal vsFileName As String)
    Dim sScriptText As String
    sScriptText = LoadStringFromFile(vsFileName)
    Call MScriptEditor1.MScripts.LoadFromText(sScriptText)
    tmrInstructionProcesor.Enabled = False
    tmrInstructionProcesor.Interval = moclMScripts.TimerInterval
    tmrInstructionProcesor.Enabled = True
    mbFileSaved = True
End Sub

Private Sub SaveMScriptsToFile(ByVal vsFileName As String)
    Dim sScriptText As String
    MScriptEditor1.SaveCurrentText
    sScriptText = MScriptEditor1.MScripts.SaveToText
    Call SaveStringToFile(sScriptText, vsFileName)
    mbFileSaved = True
End Sub

Private Sub AskAndSaveIfDirty()
    
    If Not mbFileSaved Then
        Dim enmUserResponse As VbMsgBoxResult
        enmUserResponse = MsgBox("Save the current tunes first?", vbYesNoCancel)
        Select Case enmUserResponse
            Case vbYes
                Call FileSave
            Case vbCancel
                err.Raise cdlCancel, , "Save before open canceled by user"
            Case vbNo
                'Do nothing
            Case Else
                err.Raise 1000, , "MsgBox type changed but code is not updated in frmMain.FileOpen"
        End Select
    End If
    
End Sub

Private Sub tlbMain_ButtonClick(ByVal Button As MSComctlLib.Button)

    On Error GoTo ERR_tlbMain_ButtonClick

    Select Case Button.Key
        Case sBUTTON_NEW
            Call FileNew
        Case sBUTTON_OPEN
            Call FileOpen
        Case sBUTTON_SAVE
            Call FileSave
        Case sBUTTON_HELP
            Call moHTMLHelp.ShowContent
    End Select
    
Exit Sub
ERR_tlbMain_ButtonClick:
    ShowError
End Sub

Private Sub tlbPlay_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    On Error GoTo ERR_tlbPlay_ButtonClick
    
    Select Case Button.Key
        Case sBUTTON_PLAY
            Call mnuPlay_Click
        Case sBUTTON_RECORD
            Call mnuRecord_Click
        Case sBUTTON_STOP
            Call mnuStop_Click
        Case sBUTTON_PLAY_ALL
            Call mnuPlayAll_Click
        Case Else
            err.Raise 1000, , "New button but no new code!!"
    End Select
    
Exit Sub
ERR_tlbPlay_ButtonClick:
    Call ShowError
End Sub

Private Function IsKeyForMIDIKeyboard(ByVal vlKeyCode As Long, ByVal vlShift As Long) As Boolean
    IsKeyForMIDIKeyboard = (Not (((TypeOf ActiveControl Is MScriptEditor) And (MScriptEditor1.FocusOnEditor Or MScriptEditor1.IsListViewItemEdited)) _
        Or (TypeOf ActiveControl Is MIDIKeyboard))) And (vlShift = 0)
End Function

Private Function IsKeyForMScriptEditor(ByVal vlKeyCode As Long, ByVal vlShift As Long) As Boolean
    IsKeyForMScriptEditor = ((TypeOf ActiveControl Is MScriptEditor) And MScriptEditor1.FocusOnEditor And (vlShift = 0))
End Function

Private Sub tmrFirstTime_Timer()
    tmrFirstTime.Enabled = False
    MsgBox "Click File > Open menu to see samples!", , "Thinking What To Do?"
End Sub

Private Sub tmrInstructionProcesor_Timer()
    On Error GoTo ERR_tmrInstructionProcesor_Timer
    oClock.GenerateClockPulse
Exit Sub
ERR_tmrInstructionProcesor_Timer:
    Call ShowError
End Sub

Private Sub PauseRecordingBeforePlayingSample()
    If oMScriptRecorder.IsRecording And (Not oMScriptRecorder.IsPaused) Then
        Call oMScriptRecorder.PauseRecording(True)
        bRecordingWasPausedWhilePlayingSample = True
    End If
End Sub

Private Sub StartRecordingAfterPlayingSample()
    If (bRecordingWasPausedWhilePlayingSample = True) And oMScriptRecorder.IsPaused Then
        oMScriptRecorder.StartRecording
    End If
    bRecordingWasPausedWhilePlayingSample = False
End Sub

Private Sub tmrSampleNote_Timer()
    On Error GoTo ERR_tmrSampleNote_Timer
    
    Call PauseRecordingBeforePlayingSample
    
    Select Case tmrSampleNote.Tag
        Case 1
            Call oMidiEng.StopNote(mlSAMPLE_NOTE1)
            Call oMidiEng.PlayNote(mlSAMPLE_NOTE2)
            tmrSampleNote.Tag = tmrSampleNote.Tag + 1
        Case 2
            Call oMidiEng.StopNote(mlSAMPLE_NOTE2)
            Call oMidiEng.PlayNote(mlSAMPLE_NOTE3)
            tmrSampleNote.Tag = tmrSampleNote.Tag + 1
        Case 3
            tmrSampleNote.Enabled = False
            Call oMidiEng.StopNote(mlSAMPLE_NOTE3)
            Call StartRecordingAfterPlayingSample
        Case Else
            err.Raise 1000, , "Unexpected sample note timer tag value: " & tmrSampleNote.Tag
    End Select
    
Exit Sub
ERR_tmrSampleNote_Timer:
    tmrSampleNote.Enabled = False
    tmrSampleNote.Tag = 1
    oMidiEng.StopAllNotes
    ShowError
    Call StartRecordingAfterPlayingSample
End Sub

Private Sub DDEHanlder()
    
    On Error GoTo ERR_DDEHanlder
    
    Const sDDETopic = "Main"
    
    'Try to set up the DDE link with previous instance
    BaseCommands.LinkTopic = App.EXEName & "|" & sDDETopic
    BaseCommands.LinkItem = "BaseCommands"
    BaseCommands.LinkMode = vbLinkManual
    If Command$ = vbNullString Then
        Call BaseCommands.LinkExecute(sDDE_COMMAND_ACTIVATE)
    Else
        If Not CommandSwitchProcessor(Command$) Then
            Call BaseCommands.LinkExecute(sDDE_COMMAND_LOAD & "(" & Command$ & ")")
        End If
    End If
    'If DDE command succesfully executed then end the second instant
    Unload Me
    End
Exit Sub
ERR_DDEHanlder:
    If err.Number = 293 Or err.Number = 282 Then '293 dde method invoked with no channel open, 282: No foreign application responded to DDE request
        'No DDE server present, so make this one DDE server
        Me.LinkTopic = sDDETopic
    ElseIf err.Number = 286 And App.PrevInstance = True Then    '286: DDE responce timeout - DDE server taking too much time
        MsgBox "Your request is being processed by already running instance of Visual Music. Please wait."
        Unload Me
        End
    Else
        Call ShowError
        If App.PrevInstance = True Then
            MsgBox "Instance of Visual Music is already in memory. Press Ctrl + Alt + Del and remove it and try again."
            Unload Me
            End
        End If
    End If
End Sub

Private Function CommandSwitchProcessor(ByVal vsCommandLine As String) As Boolean
    Select Case LCase(Trim(vsCommandLine))
        Case "/regicon"
            Call RegisterFileExtensionAndIcon("MSC", "MScript File", GetPathWithSlash(App.Path) & App.EXEName & ".exe", GetPathWithSlash(App.Path) & "VisualMusic.ico")
            Call SaveSetting(msREG_APP_ROOT, msREG_SECTION_STATE, msREG_KEY_EXT_REGISTERED, "True")
            Unload Me
            End
        Case Else
            CommandSwitchProcessor = False
    End Select
End Function

Private Sub FillSampleScript()

    Dim sSampleScript As String
    sSampleScript = sSampleScript & "-----------Sample Script-------------" & vbCrLf & vbCrLf
    sSampleScript = sSampleScript & "LABEL:DoAgain" & vbCrLf & vbCrLf
    sSampleScript = sSampleScript & vbTab & "'This is comment - Play musical notes" & vbCrLf & vbCrLf
    sSampleScript = sSampleScript & vbTab & "a a# c c#" & vbCrLf & vbCrLf
    sSampleScript = sSampleScript & "JUMP:DoAgain" & vbCrLf & vbCrLf
    sSampleScript = sSampleScript & "-----Press Play button to run the script" & vbCrLf & vbCrLf
    
    moclMScripts("Tune1").Text = sSampleScript
    moclMScripts("Tune2").Text = "--Write script for Tune2 here" & vbCrLf & vbCrLf & "--You can run all the tunes simultaneously!"
    moclMScripts("Tune3").Text = "--To change the instrument write like this" & vbCrLf & "INSTRUMENT:Chorus_Piano" & vbCrLf & vbCrLf & vbCrLf & "--Press Help for more info!"
    'Don't prompt for saving sample tune
    mbFileSaved = True
End Sub
