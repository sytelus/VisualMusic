VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl MIDIParamControls 
   ClientHeight    =   2610
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4290
   LockControls    =   -1  'True
   ScaleHeight     =   2610
   ScaleWidth      =   4290
   Begin MSComctlLib.Slider sldVolume 
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   1320
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   10
      Max             =   100
      TickStyle       =   3
   End
   Begin MSComctlLib.Toolbar tlbOctave 
      Height          =   570
      Left            =   60
      TabIndex        =   1
      Top             =   240
      Width           =   4170
      _ExtentX        =   7355
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&1"
            Style           =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&2"
            Style           =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   2
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   2
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   2
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   2
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   2
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&C"
            Style           =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.Slider sldPan 
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   2100
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   10
      Max             =   100
      TickStyle       =   3
   End
   Begin VB.Label lblVolumeDisplay 
      AutoSize        =   -1  'True
      Caption         =   "XXX"
      Height          =   195
      Left            =   660
      TabIndex        =   8
      Top             =   1080
      Width           =   315
   End
   Begin VB.Label lblPanDisplay 
      AutoSize        =   -1  'True
      Caption         =   "XXX"
      Height          =   195
      Left            =   420
      TabIndex        =   7
      Top             =   1860
      Width           =   315
   End
   Begin VB.Label lblOctaveDisplay 
      AutoSize        =   -1  'True
      Caption         =   "XXX"
      Height          =   195
      Left            =   660
      TabIndex        =   6
      Top             =   0
      Width           =   315
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "&Pan:"
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   1860
      Width           =   330
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Volume:"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   1080
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Octave:"
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   570
   End
End
Attribute VB_Name = "MIDIParamControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_lOctave As Long
Private m_lVolume As Long
Private m_lPan As Long

Public Event SetOctave(ByRef rlOctave As Long)
Public Event SetVolume(ByRef rlVolume As Long)
Public Event SetPan(ByRef rlPan As Long)

Private Sub sldPan_Change()
    RaiseEvent SetPan(sldPan.Value)
    Pan = sldPan.Value
End Sub

Private Sub sldVolume_Change()
    RaiseEvent SetVolume(sldVolume.Value)
    Volume = sldVolume.Value
End Sub

Private Sub tlbOctave_ButtonClick(ByVal Button As MSComctlLib.Button)
    RaiseEvent SetOctave(Button.Index)
    m_lOctave = Button.Index
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Dim lButtonIndex As Long
    
    For lButtonIndex = 1 To tlbOctave.Buttons.Count
        If tlbOctave.Buttons(lButtonIndex).Index < 10 Then
            tlbOctave.Buttons(lButtonIndex).Caption = "&" & tlbOctave.Buttons(lButtonIndex).Index
        Else
            tlbOctave.Buttons(lButtonIndex).Caption = "&" & Hex(tlbOctave.Buttons(lButtonIndex).Index)
        End If
        tlbOctave.Buttons(lButtonIndex).Style = tbrButtonGroup
    Next lButtonIndex
    Octave = 1
    Volume = 0
    Pan = 0
End Sub

Public Property Get Octave() As Long
    Octave = m_lOctave
End Property

Public Property Get Volume() As Long
    Volume = m_lVolume
End Property

Public Property Get Pan() As Long
    Pan = m_lPan
End Property

Public Property Let Octave(ByVal vlOctave As Long)
    m_lOctave = vlOctave
    tlbOctave.Buttons(vlOctave).Value = tbrPressed
    lblOctaveDisplay.Caption = m_lOctave
End Property

Public Property Let Volume(ByVal vlVolume As Long)
    m_lVolume = vlVolume
    sldVolume.Value = vlVolume
    lblVolumeDisplay.Caption = m_lVolume
End Property

Public Property Let Pan(ByVal vlPan As Long)
    m_lPan = vlPan
    sldPan.Value = vlPan
    lblPanDisplay.Caption = m_lPan
End Property

Public Function ProcessKey(ByVal KeyCode As Long, ByVal Shift As Long) As Boolean

    Dim lPressedButtonIndex As Long

    ProcessKey = True
    If Shift And vbAltMask Then
        If (KeyCode >= vbKey1 And KeyCode <= vbKey9) Then
            tlbOctave.Buttons(KeyCode - vbKey1 + 1).Value = tbrPressed
            Call tlbOctave_ButtonClick(tlbOctave.Buttons(KeyCode - vbKey1 + 1))
        ElseIf (KeyCode = vbKeyA) Or (KeyCode = vbKeyB) Or (KeyCode = vbKeyC) Then
            tlbOctave.Buttons(KeyCode - vbKeyA + 10).Value = tbrUnpressed
            Call tlbOctave_ButtonClick(tlbOctave.Buttons(KeyCode - vbKeyA + 10))
        ElseIf KeyCode = vbKeyV Then
            sldVolume.SetFocus
        ElseIf KeyCode = vbKeyP Then
            sldPan.SetFocus
        Else
            ProcessKey = False
        End If
    Else
        ProcessKey = False
    End If
End Function

Private Function GetPressedButtonIndex() As Long
    Dim lButtonIndex As Long
    
    GetPressedButtonIndex = -1
    
    For lButtonIndex = 1 To tlbOctave.Buttons.Count
        If tlbOctave.Buttons(lButtonIndex).Value = tbrPressed Then
            GetPressedButtonIndex = lButtonIndex
            Exit For
        End If
    Next lButtonIndex
    
End Function

