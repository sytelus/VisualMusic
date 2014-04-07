VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Visual Music"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4260
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   4560
   ScaleWidth      =   4260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Height          =   75
      Left            =   120
      TabIndex        =   14
      Top             =   3720
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   435
      Left            =   1560
      TabIndex        =   13
      Top             =   3960
      Width           =   1155
   End
   Begin VB.Frame Frame2 
      Height          =   75
      Left            =   120
      TabIndex        =   9
      Top             =   2700
      Width           =   4095
   End
   Begin VB.Frame Frame1 
      Height          =   75
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   4095
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   180
      Picture         =   "About.frx":1CFA
      ScaleHeight     =   480
      ScaleWidth      =   420
      TabIndex        =   0
      Top             =   180
      Width           =   420
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Visual Music HomePage"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   12
      Top             =   2400
      Width           =   1725
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "This program is freeware. Source code available upon request. "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   120
      TabIndex        =   11
      Top             =   3180
      Width           =   4575
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "© Shital Shah, 25 July 1999."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   10
      Top             =   2940
      Width           =   4335
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "http://i.am/vmusic"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   1
      Left            =   1980
      MouseIcon       =   "About.frx":2072
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   2400
      Width           =   1320
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "http://i.am/shital"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   1980
      MouseIcon       =   "About.frx":237C
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   2100
      Width           =   1215
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      Caption         =   "Shital Shah,"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2100
      Width           =   915
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      Caption         =   "Visual Music is designed and developed by,"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1740
      Width           =   4335
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Caption         =   $"About.frx":2686
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   4035
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      TabIndex        =   2
      Top             =   360
      Width           =   60
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Visual Music"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   780
      TabIndex        =   1
      Top             =   360
      Width           =   1440
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    lblVersion.Caption = App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Label5_Click(Index As Integer)
    Call OpenAnyFile(Label5(Index).Caption)
End Sub
