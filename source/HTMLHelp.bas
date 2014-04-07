Attribute VB_Name = "modHTMLHelp"
Option Explicit

Private Declare Function HTMLHelp Lib "hhctrl.ocx" Alias "HTMLHelpA(Byval hWndCaller as long, Byval pszFile as string, byval uCommand as long, byval dwData as long) as long" ()
Private Const HH_DISPLAY_TOPIC As Long = 0
HH_DISPLAY_TOC = 1
HH_DISPLAY_SEARCH = 3

