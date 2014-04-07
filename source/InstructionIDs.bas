Attribute VB_Name = "InstructionProcessorGlobals"
Option Explicit

Public Enum InstructionProcessorStatus
    ipsIdle = 0
    ipsPlayingNote = 1
    ipsPlayingSilence = 2
    ipsExecutingCommand = 3
End Enum

Public Const glINSTRUCTION_PLAY_NOTE As Long = 1
Public Const glINSTRUCTION_SET_INSTRUMENT As Long = 2
Public Const glINSTRUCTION_SET_OCTAVE As Long = 3
Public Const glINSTRUCTION_SET_VOLUME As Long = 4
Public Const glINSTRUCTION_SET_PAN As Long = 5
Public Const glINSTRUCTION_PLAY_SILENCE As Long = 6
Public Const glINSTRUCTION_JUMP As Long = 7
Public Const glINSTRUCTION_LOOP As Long = 8

Public Const glARRAY_INDEX_INSTRUCTION_CODE As Long = 0
Public Const glARRAY_INDEX_FIRST_PARAM As Long = 1

Public Const glPLAY_NOTE_PARAM_NOTE As Long = 1
Public Const glPLAY_NOTE_PARAM_NOTE_INTERVAL As Long = 2
Public Const glPLAY_NOTE_PARAM_SILENCE_INTERVAL As Long = 3
Public Const glPLAY_NOTE_PARAM_OCTAVE As Long = 4
Public Const glPLAY_NOTE_PARAM_VOLUME As Long = 5
Public Const glPLAY_NOTE_PARAM_PAN As Long = 6
Public Const glPLAY_NOTE_PARAM_INSTRUMENT As Long = 7

Public Const glLOOP_PARAM_NAME As Long = 1
Public Const glLOOP_PARAM_STOP As Long = 2
Public Const glLOOP_PARAM_START As Long = 3
Public Const glLOOP_PARAM_INCREMENT As Long = 4

