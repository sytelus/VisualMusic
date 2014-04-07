Attribute VB_Name = "modInstrumentManager"
Option Explicit

Public Enum InstrumentType
    'Do not change number assignments - assumed by InstrumentManager.GetInstrumentType
    itpPiano = 0
    itpChromaticPercussion = 1
    itpOrgan = 2
    itpGuitar = 3
    itpBass = 4
    itpStrings = 5
    itpEnsemble = 6
    itpBrass = 7
    itpReed = 8
    itpPipe = 9
    itpSynthLead = 10
    itpMisc1 = 11
    itpFX = 12
    itpMisc2 = 13
    itpSynthPad = 14
    itpSoundEffects = 15
End Enum

Public Const gsDEFAULT_INSTRUMENTS_INI As String = "instruments.ini"
Public Const gsDEFAULT_FAV_INSTRUMENTS_INI As String = "favorites.ini"

