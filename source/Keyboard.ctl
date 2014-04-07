VERSION 5.00
Begin VB.UserControl MIDIKeyboard 
   ClientHeight    =   1215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11040
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   1215
   ScaleWidth      =   11040
   Begin VB.PictureBox pnlKeyboard 
      AutoSize        =   -1  'True
      ForeColor       =   &H00808080&
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   9255
      TabIndex        =   0
      Top             =   0
      Width           =   9315
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         DragIcon        =   "Keyboard.ctx":0000
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   585
         Index           =   54
         Left            =   7680
         ScaleHeight     =   525
         ScaleWidth      =   105
         TabIndex        =   31
         Top             =   90
         Width           =   165
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         DragIcon        =   "Keyboard.ctx":030A
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   585
         Index           =   42
         Left            =   6000
         ScaleHeight     =   525
         ScaleWidth      =   105
         TabIndex        =   30
         Top             =   90
         Width           =   165
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         DragIcon        =   "Keyboard.ctx":0614
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   585
         Index           =   30
         Left            =   4320
         ScaleHeight     =   525
         ScaleWidth      =   105
         TabIndex        =   29
         Top             =   90
         Width           =   165
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         DragIcon        =   "Keyboard.ctx":091E
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   585
         Index           =   18
         Left            =   2640
         ScaleHeight     =   525
         ScaleWidth      =   105
         TabIndex        =   28
         Top             =   90
         Width           =   165
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         DragIcon        =   "Keyboard.ctx":0C28
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   585
         Index           =   56
         Left            =   7920
         ScaleHeight     =   525
         ScaleWidth      =   105
         TabIndex        =   27
         Top             =   90
         Width           =   165
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         DragIcon        =   "Keyboard.ctx":0F32
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   585
         Index           =   51
         Left            =   7200
         ScaleHeight     =   525
         ScaleWidth      =   105
         TabIndex        =   26
         Top             =   90
         Width           =   165
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         DragIcon        =   "Keyboard.ctx":123C
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   585
         Index           =   49
         Left            =   6960
         ScaleHeight     =   525
         ScaleWidth      =   105
         TabIndex        =   25
         Top             =   90
         Width           =   165
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         DragIcon        =   "Keyboard.ctx":1546
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   585
         Index           =   46
         Left            =   6480
         ScaleHeight     =   525
         ScaleWidth      =   105
         TabIndex        =   24
         Top             =   90
         Width           =   165
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         DragIcon        =   "Keyboard.ctx":1850
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   585
         Index           =   44
         Left            =   6240
         ScaleHeight     =   525
         ScaleWidth      =   105
         TabIndex        =   23
         Top             =   90
         Width           =   165
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         DragIcon        =   "Keyboard.ctx":1B5A
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   585
         Index           =   39
         Left            =   5520
         ScaleHeight     =   525
         ScaleWidth      =   105
         TabIndex        =   22
         Top             =   90
         Width           =   165
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         DragIcon        =   "Keyboard.ctx":1E64
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   585
         Index           =   37
         Left            =   5280
         ScaleHeight     =   525
         ScaleWidth      =   105
         TabIndex        =   21
         Top             =   90
         Width           =   165
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         DragIcon        =   "Keyboard.ctx":216E
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   585
         Index           =   34
         Left            =   4800
         ScaleHeight     =   525
         ScaleWidth      =   105
         TabIndex        =   20
         Top             =   90
         Width           =   165
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         DragIcon        =   "Keyboard.ctx":2478
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   585
         Index           =   32
         Left            =   4560
         ScaleHeight     =   525
         ScaleWidth      =   105
         TabIndex        =   19
         Top             =   90
         Width           =   165
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         DragIcon        =   "Keyboard.ctx":2782
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   585
         Index           =   27
         Left            =   3840
         ScaleHeight     =   525
         ScaleWidth      =   105
         TabIndex        =   18
         Top             =   90
         Width           =   165
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         DragIcon        =   "Keyboard.ctx":2A8C
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   585
         Index           =   25
         Left            =   3600
         ScaleHeight     =   525
         ScaleWidth      =   105
         TabIndex        =   17
         Top             =   90
         Width           =   165
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         DragIcon        =   "Keyboard.ctx":2D96
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   585
         Index           =   22
         Left            =   3120
         ScaleHeight     =   525
         ScaleWidth      =   105
         TabIndex        =   16
         Top             =   90
         Width           =   165
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         DragIcon        =   "Keyboard.ctx":30A0
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   585
         Index           =   20
         Left            =   2880
         ScaleHeight     =   525
         ScaleWidth      =   105
         TabIndex        =   15
         Top             =   90
         Width           =   165
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         DragIcon        =   "Keyboard.ctx":33AA
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   585
         Index           =   15
         Left            =   2160
         ScaleHeight     =   525
         ScaleWidth      =   105
         TabIndex        =   14
         Top             =   90
         Width           =   165
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         DragIcon        =   "Keyboard.ctx":36B4
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   585
         Index           =   13
         Left            =   1920
         ScaleHeight     =   525
         ScaleWidth      =   105
         TabIndex        =   13
         Top             =   90
         Width           =   165
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         DragIcon        =   "Keyboard.ctx":39BE
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   585
         Index           =   10
         Left            =   1440
         ScaleHeight     =   525
         ScaleWidth      =   105
         TabIndex        =   12
         Top             =   90
         Width           =   165
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         DragIcon        =   "Keyboard.ctx":3CC8
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   585
         Index           =   8
         Left            =   1200
         ScaleHeight     =   525
         ScaleWidth      =   105
         TabIndex        =   11
         Top             =   90
         Width           =   165
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         DragIcon        =   "Keyboard.ctx":3FD2
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   585
         Index           =   6
         Left            =   960
         ScaleHeight     =   525
         ScaleWidth      =   105
         TabIndex        =   10
         Top             =   90
         Width           =   165
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         DragIcon        =   "Keyboard.ctx":42DC
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   585
         Index           =   3
         Left            =   480
         ScaleHeight     =   525
         ScaleWidth      =   105
         TabIndex        =   9
         Top             =   90
         Width           =   165
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         DragIcon        =   "Keyboard.ctx":45E6
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   585
         Index           =   1
         Left            =   240
         ScaleHeight     =   525
         ScaleWidth      =   105
         TabIndex        =   8
         Top             =   90
         Width           =   165
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         DragIcon        =   "Keyboard.ctx":48F0
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   585
         Index           =   61
         Left            =   8640
         ScaleHeight     =   525
         ScaleWidth      =   105
         TabIndex        =   7
         Top             =   90
         Width           =   165
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         DragIcon        =   "Keyboard.ctx":4BFA
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   585
         Index           =   63
         Left            =   8910
         ScaleHeight     =   525
         ScaleWidth      =   105
         TabIndex        =   6
         Top             =   90
         Width           =   165
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         DragIcon        =   "Keyboard.ctx":4F04
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   585
         Index           =   58
         Left            =   8160
         ScaleHeight     =   525
         ScaleWidth      =   105
         TabIndex        =   5
         Top             =   90
         Width           =   165
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         DragIcon        =   "Keyboard.ctx":520E
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   855
         Index           =   62
         Left            =   8730
         ScaleHeight     =   795
         ScaleWidth      =   195
         TabIndex        =   3
         Top             =   90
         Width           =   255
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         DragIcon        =   "Keyboard.ctx":5518
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   855
         Index           =   64
         Left            =   8970
         ScaleHeight     =   795
         ScaleWidth      =   195
         TabIndex        =   2
         Top             =   90
         Width           =   255
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         DragIcon        =   "Keyboard.ctx":5822
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   855
         Index           =   7
         Left            =   1050
         ScaleHeight     =   795
         ScaleWidth      =   195
         TabIndex        =   32
         Top             =   90
         Width           =   255
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         DragIcon        =   "Keyboard.ctx":5B2C
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   855
         Index           =   9
         Left            =   1290
         ScaleHeight     =   795
         ScaleWidth      =   195
         TabIndex        =   33
         Top             =   90
         Width           =   255
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         DragIcon        =   "Keyboard.ctx":5E36
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   855
         Index           =   11
         Left            =   1530
         ScaleHeight     =   795
         ScaleWidth      =   195
         TabIndex        =   34
         Top             =   90
         Width           =   255
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         DragIcon        =   "Keyboard.ctx":6140
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   855
         Index           =   4
         Left            =   570
         ScaleHeight     =   795
         ScaleWidth      =   195
         TabIndex        =   35
         Top             =   90
         Width           =   255
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         DragIcon        =   "Keyboard.ctx":644A
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   855
         Index           =   5
         Left            =   810
         ScaleHeight     =   795
         ScaleWidth      =   195
         TabIndex        =   36
         Top             =   90
         Width           =   255
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         DragIcon        =   "Keyboard.ctx":6754
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   855
         Index           =   2
         Left            =   330
         ScaleHeight     =   795
         ScaleWidth      =   195
         TabIndex        =   37
         Top             =   90
         Width           =   255
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         DragIcon        =   "Keyboard.ctx":6A5E
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   855
         Index           =   12
         Left            =   1770
         ScaleHeight     =   795
         ScaleWidth      =   195
         TabIndex        =   38
         Top             =   90
         Width           =   255
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         DragIcon        =   "Keyboard.ctx":6D68
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   855
         Index           =   14
         Left            =   2010
         ScaleHeight     =   795
         ScaleWidth      =   195
         TabIndex        =   39
         Top             =   90
         Width           =   255
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         DragIcon        =   "Keyboard.ctx":7072
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   855
         Index           =   16
         Left            =   2250
         ScaleHeight     =   795
         ScaleWidth      =   195
         TabIndex        =   40
         Top             =   90
         Width           =   255
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         DragIcon        =   "Keyboard.ctx":737C
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   855
         Index           =   17
         Left            =   2490
         ScaleHeight     =   795
         ScaleWidth      =   195
         TabIndex        =   41
         Top             =   90
         Width           =   255
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         DragIcon        =   "Keyboard.ctx":7686
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   855
         Index           =   19
         Left            =   2730
         ScaleHeight     =   795
         ScaleWidth      =   195
         TabIndex        =   42
         Top             =   90
         Width           =   255
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         DragIcon        =   "Keyboard.ctx":7990
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   855
         Index           =   21
         Left            =   2970
         ScaleHeight     =   795
         ScaleWidth      =   195
         TabIndex        =   43
         Top             =   90
         Width           =   255
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         DragIcon        =   "Keyboard.ctx":7C9A
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   855
         Index           =   23
         Left            =   3210
         ScaleHeight     =   795
         ScaleWidth      =   195
         TabIndex        =   44
         Top             =   90
         Width           =   255
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         DragIcon        =   "Keyboard.ctx":7FA4
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   855
         Index           =   24
         Left            =   3450
         ScaleHeight     =   795
         ScaleWidth      =   195
         TabIndex        =   45
         Top             =   90
         Width           =   255
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         DragIcon        =   "Keyboard.ctx":82AE
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   855
         Index           =   26
         Left            =   3690
         ScaleHeight     =   795
         ScaleWidth      =   195
         TabIndex        =   46
         Top             =   90
         Width           =   255
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         DragIcon        =   "Keyboard.ctx":85B8
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   855
         Index           =   28
         Left            =   3930
         ScaleHeight     =   795
         ScaleWidth      =   195
         TabIndex        =   47
         Top             =   90
         Width           =   255
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         DragIcon        =   "Keyboard.ctx":88C2
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   855
         Index           =   29
         Left            =   4170
         ScaleHeight     =   795
         ScaleWidth      =   195
         TabIndex        =   48
         Top             =   90
         Width           =   255
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         DragIcon        =   "Keyboard.ctx":8BCC
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   855
         Index           =   0
         Left            =   90
         ScaleHeight     =   795
         ScaleWidth      =   195
         TabIndex        =   65
         Top             =   90
         Width           =   255
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         DragIcon        =   "Keyboard.ctx":8ED6
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   855
         Index           =   59
         Left            =   8250
         ScaleHeight     =   795
         ScaleWidth      =   195
         TabIndex        =   1
         Top             =   90
         Width           =   255
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         DragIcon        =   "Keyboard.ctx":91E0
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   855
         Index           =   60
         Left            =   8490
         ScaleHeight     =   795
         ScaleWidth      =   195
         TabIndex        =   4
         Top             =   90
         Width           =   255
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         DragIcon        =   "Keyboard.ctx":94EA
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   855
         Index           =   31
         Left            =   4410
         ScaleHeight     =   795
         ScaleWidth      =   195
         TabIndex        =   49
         Top             =   90
         Width           =   255
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         DragIcon        =   "Keyboard.ctx":97F4
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   855
         Index           =   33
         Left            =   4650
         ScaleHeight     =   795
         ScaleWidth      =   195
         TabIndex        =   50
         Top             =   90
         Width           =   255
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         DragIcon        =   "Keyboard.ctx":9AFE
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   855
         Index           =   35
         Left            =   4890
         ScaleHeight     =   795
         ScaleWidth      =   195
         TabIndex        =   51
         Top             =   90
         Width           =   255
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         DragIcon        =   "Keyboard.ctx":9E08
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   855
         Index           =   36
         Left            =   5130
         ScaleHeight     =   795
         ScaleWidth      =   195
         TabIndex        =   52
         Top             =   90
         Width           =   255
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         DragIcon        =   "Keyboard.ctx":A112
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   855
         Index           =   38
         Left            =   5370
         ScaleHeight     =   795
         ScaleWidth      =   195
         TabIndex        =   53
         Top             =   90
         Width           =   255
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         DragIcon        =   "Keyboard.ctx":A41C
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   855
         Index           =   40
         Left            =   5610
         ScaleHeight     =   795
         ScaleWidth      =   195
         TabIndex        =   54
         Top             =   90
         Width           =   255
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         DragIcon        =   "Keyboard.ctx":A726
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   855
         Index           =   41
         Left            =   5850
         ScaleHeight     =   795
         ScaleWidth      =   195
         TabIndex        =   55
         Top             =   90
         Width           =   255
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         DragIcon        =   "Keyboard.ctx":AA30
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   855
         Index           =   43
         Left            =   6090
         ScaleHeight     =   795
         ScaleWidth      =   195
         TabIndex        =   56
         Top             =   90
         Width           =   255
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         DragIcon        =   "Keyboard.ctx":AD3A
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   855
         Index           =   45
         Left            =   6330
         ScaleHeight     =   795
         ScaleWidth      =   195
         TabIndex        =   57
         Top             =   90
         Width           =   255
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         DragIcon        =   "Keyboard.ctx":B044
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   855
         Index           =   47
         Left            =   6570
         ScaleHeight     =   795
         ScaleWidth      =   195
         TabIndex        =   58
         Top             =   90
         Width           =   255
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         DragIcon        =   "Keyboard.ctx":B34E
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   855
         Index           =   48
         Left            =   6810
         ScaleHeight     =   795
         ScaleWidth      =   195
         TabIndex        =   59
         Top             =   90
         Width           =   255
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         DragIcon        =   "Keyboard.ctx":B658
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   855
         Index           =   50
         Left            =   7050
         ScaleHeight     =   795
         ScaleWidth      =   195
         TabIndex        =   60
         Top             =   90
         Width           =   255
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         DragIcon        =   "Keyboard.ctx":B962
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   855
         Index           =   52
         Left            =   7290
         ScaleHeight     =   795
         ScaleWidth      =   195
         TabIndex        =   61
         Top             =   90
         Width           =   255
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         DragIcon        =   "Keyboard.ctx":BC6C
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   855
         Index           =   53
         Left            =   7530
         ScaleHeight     =   795
         ScaleWidth      =   195
         TabIndex        =   62
         Top             =   90
         Width           =   255
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         DragIcon        =   "Keyboard.ctx":BF76
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   855
         Index           =   55
         Left            =   7770
         ScaleHeight     =   795
         ScaleWidth      =   195
         TabIndex        =   63
         Top             =   90
         Width           =   255
      End
      Begin VB.PictureBox PanelWhite 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         DragIcon        =   "Keyboard.ctx":C280
         DragMode        =   1  'Automatic
         ForeColor       =   &H00808080&
         Height          =   855
         Index           =   57
         Left            =   8010
         ScaleHeight     =   795
         ScaleWidth      =   195
         TabIndex        =   64
         Top             =   90
         Width           =   255
      End
      Begin VB.Shape shpFocusIndicator 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H000000FF&
         FillStyle       =   2  'Horizontal Line
         Height          =   90
         Left            =   90
         Shape           =   3  'Circle
         Top             =   0
         Width           =   90
      End
   End
End
Attribute VB_Name = "MIDIKeyboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event KeyDown(ByVal vlKeyIndex As Long)
Public Event KeyUp(ByVal vlKeyIndex As Long)

Private mbIsKeyDown As Boolean
Private mlLastKeyDown As Long
Private m_bActive As Boolean
Private m_bKeyUpDownEffect As Boolean
Private m_oclKeyBoardMask As Collection
Private m_bShowFocusIndicator As Boolean

Public Property Get ShowFocusIndicator() As Boolean
    ShowFocusIndicator = m_bShowFocusIndicator
End Property

Public Property Let ShowFocusIndicator(ByVal vboolShowFocusIndicator As Boolean)
    m_bShowFocusIndicator = vboolShowFocusIndicator
    shpFocusIndicator.Visible = m_bShowFocusIndicator
    Call PropertyChanged("ShowFocusIndicator")
End Property

Private Sub PanelWhite_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
    'Call MakeKeyUp(Index)
    Call MakeAllKeyDown
End Sub

Public Sub MakeAllKeyDown()
    Dim lKeyIndex As Long
    
    For lKeyIndex = PanelWhite.LBound To PanelWhite.UBound
        If IsMajorNote(lKeyIndex) Then
            If PanelWhite(lKeyIndex).BackColor = vbBlack Then
                Call MakeKeyUp(lKeyIndex)
            End If
        Else
            If PanelWhite(lKeyIndex).BackColor = vbWhite Then
                Call MakeKeyUp(lKeyIndex)
            End If
        End If
    Next lKeyIndex
End Sub

Private Sub PanelWhite_DragOver(Index As Integer, Source As Control, x As Single, y As Single, State As Integer)
    Call MakeKeyDown(Index)
    UserControl.SetFocus
End Sub

Public Function IsKeyDown() As Boolean
    IsKeyDown = mbIsKeyDown
End Function

Private Sub UserControl_EnterFocus()
    shpFocusIndicator.BackColor = vbGreen
End Sub

Private Sub UserControl_ExitFocus()
    shpFocusIndicator.BackColor = vbRed
End Sub

Private Sub UserControl_Initialize()
    mbIsKeyDown = False
    mlLastKeyDown = -1
    Set m_oclKeyBoardMask = New Collection
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 Then
        Call MakeKeyDown(KeyCode, True)
    End If
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    If Shift = 0 Then
        Call MakeKeyUp(KeyCode, True)
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_bActive = PropBag.ReadProperty("Active", True)
    m_bKeyUpDownEffect = PropBag.ReadProperty("KeyUpDownEffect", True)
    m_bShowFocusIndicator = PropBag.ReadProperty("ShowFocusIndicator", True)
    shpFocusIndicator.Visible = m_bShowFocusIndicator
End Sub

Private Sub UserControl_Resize()
    pnlKeyboard.Left = 0
    pnlKeyboard.Top = 0
    Height = pnlKeyboard.Height
    Width = pnlKeyboard.Width
End Sub

Public Property Get Active() As Boolean
    Active = m_bActive
End Property

Public Property Let Active(ByVal vlActiveStatus As Boolean)
    m_bActive = vlActiveStatus
    Call PropertyChanged("Active")
End Property

Private Sub UserControl_Terminate()
    Set m_oclKeyBoardMask = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Active", m_bActive, True
    PropBag.WriteProperty "KeyUpDownEffect", m_bKeyUpDownEffect, True
    PropBag.WriteProperty "ShowFocusIndicator", m_bShowFocusIndicator, True
End Sub

Public Sub MakeKeyDown(ByVal lIndexOrKeyCode As Long, Optional ByVal vboolIsKeyCode As Boolean = False, Optional ByVal vboolVisualEffectOnly As Boolean = False)
    
    Dim lIndex As Long
    If vboolIsKeyCode = True Then
        lIndex = GetIndexForKeyCode(lIndexOrKeyCode)
    Else
        lIndex = lIndexOrKeyCode
    End If
    
    If lIndex <> -1 Then
        If m_bActive = True Then
            If Not (mbIsKeyDown And mlLastKeyDown = lIndex) Then
                mbIsKeyDown = True
                mlLastKeyDown = lIndex
                If m_bKeyUpDownEffect Then
                    Call SetKeyColor(lIndex, True)
                End If
                If Not vboolVisualEffectOnly Then
                    RaiseEvent KeyDown(lIndex)
                End If
            End If
        End If
    End If
End Sub

Public Sub MakeKeyUp(ByVal lIndexOrKeyCode As Long, Optional ByVal vboolIsKeyCode As Boolean = False, Optional ByVal vboolVisualEffectOnly As Boolean = False)

    Dim lIndex As Long
    If vboolIsKeyCode = True Then
        lIndex = GetIndexForKeyCode(lIndexOrKeyCode)
    Else
        lIndex = lIndexOrKeyCode
    End If

    If lIndex <> -1 Then
        If m_bActive = True Then
            mbIsKeyDown = False
            If m_bKeyUpDownEffect Then
                Call SetKeyColor(lIndex, False)
            End If
            If Not vboolVisualEffectOnly Then
                RaiseEvent KeyUp(lIndex)
            End If
        End If
    End If
End Sub

Public Sub ClearKeyboardLayout()
    Call ClearCollection(m_oclKeyBoardMask)
End Sub

Public Sub SetKeyCodeForIndex(ByVal venmKeyCode As VBRUN.KeyCodeConstants, ByVal vlIndexForKeyCode As Long)
    If IsItemExistInCol(m_oclKeyBoardMask, venmKeyCode) Then
        m_oclKeyBoardMask(venmKeyCode) = vlIndexForKeyCode
    Else
        Call m_oclKeyBoardMask.Add(CStr(vlIndexForKeyCode), CStr(venmKeyCode))
    End If
End Sub

Public Function GetIndexForKeyCode(ByVal venmKeyCode As VBRUN.KeyCodeConstants) As Long
    If IsItemExistInCol(m_oclKeyBoardMask, venmKeyCode) Then
        GetIndexForKeyCode = m_oclKeyBoardMask(CStr(venmKeyCode))
    Else
        GetIndexForKeyCode = -1
    End If
End Function

Public Property Get KeyUpDownEffect() As Boolean
    KeyUpDownEffect = m_bKeyUpDownEffect
End Property

Public Property Let KeyUpDownEffect(ByVal vboolShowEffect As Boolean)
    m_bKeyUpDownEffect = vboolShowEffect
    Call PropertyChanged("Active")
End Property

Private Sub SetKeyColor(ByVal vlKeyIndex As Long, ByVal vboolIsPressedState As Boolean)
    If (vboolIsPressedState And IsMajorNote(vlKeyIndex)) _
        Or ((Not vboolIsPressedState) And (Not IsMajorNote(vlKeyIndex))) Then
          PanelWhite(vlKeyIndex).BackColor = vbBlack
    Else
          PanelWhite(vlKeyIndex).BackColor = vbWhite
    End If
End Sub

Private Function IsMajorNote(ByVal vlKeyIndex As Long) As Boolean
    Select Case vlKeyIndex
        Case 1, 3, 6, 8, 10, 13, 15, 18, 20, 22, 25, 27, 30, 32, 34, 37, 39, 42, 44, 46, 49, 51, 54, 56, 58, 61, 63
            IsMajorNote = False
        Case Else
            IsMajorNote = True
    End Select
End Function
