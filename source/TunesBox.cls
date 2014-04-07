VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TunesBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_oclMScriptTexts As MScriptTexts

Public Property Get MScriptTexts() As MScriptTexts
    Set MScriptTexts = m_oclMScriptTexts
End Property

Public Property Set MScriptTexts(ByVal voclMScriptTexts As MScriptTexts)
    Set m_oclMScriptTexts = voclMScriptTexts
End Property

Public Sub CompileMScripts(Optional ByVal vvaScriptNames As Variant = Empty)
    If IsEmpty(vvaScriptNames) Then
        'Compile all
        
    Else
        Err.Raise 1000, , "Not implemented"
    End If
End Sub

