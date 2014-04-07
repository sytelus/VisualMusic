Attribute VB_Name = "modErrorConst"
Option Explicit

Public Const lERR_BASE As Long = vbObjectError + 512

Public Const lERR_BASE_COMPILE As Long = lERR_BASE + 10000
Public Const lERR_COMPILE_HIGHEST As Long = lERR_BASE + 11000

Public Const lERR_BASE_RUN As Long = lERR_BASE + 20000

Public Const lERR_BASE_MIDIENG As Long = lERR_BASE + 30000
Public Const lERR_MIDIENG_HIGHEST As Long = lERR_BASE + 31000

'Compilation errors
Public Const lERR_COMPILE_MISSING_PARAM As Long = lERR_BASE_COMPILE + 1
Public Const lERR_COMPILE_MORE_PARAM As Long = lERR_BASE_COMPILE + 2
Public Const lERR_COMPILE_NO_LABEL_NAME As Long = lERR_BASE_COMPILE + 3
Public Const lERR_COMPILE_INVALID_INSTRUCTION As Long = lERR_BASE_COMPILE + 4
Public Const lERR_COMPILE_INVALID_INSTRUMENT As Long = lERR_BASE_COMPILE + 5
Public Const lERR_COMPILE_NO_INSTRUMENT_NAME_SUPPORT As Long = lERR_BASE_COMPILE + 6
Public Const lERR_COMPILE_LABEL_ALREADY_DECLARED As Long = lERR_BASE_COMPILE + 7

'Runtime errors
Public Const lERR_RUN_INVALID_VAR_NAME As Long = lERR_BASE_RUN + 1
Public Const lERR_RUN_INVALID_LABLE_NAME As Long = lERR_BASE_RUN + 2


