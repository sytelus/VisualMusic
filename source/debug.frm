VERSION 5.00
Begin VB.Form frmDebug 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Debug Window"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "debug.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   315
      Left            =   3480
      TabIndex        =   1
      Top             =   3240
      Width           =   1155
   End
   Begin VB.TextBox txtDebug 
      BackColor       =   &H8000000F&
      Height          =   3135
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   4635
   End
End
Attribute VB_Name = "frmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClear_Click()
    txtDebug.Text = vbNullString
End Sub
