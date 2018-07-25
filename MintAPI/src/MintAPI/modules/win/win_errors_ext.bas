Attribute VB_Name = "win_errors_ext"
Option Explicit

  
Public Const STATUS_WAIT_0                    As Long = &H0
Public Const STATUS_ABANDONED_WAIT_0          As Long = &H80
Public Const STATUS_USER_APC                  As Long = &HC0
Public Const STATUS_TIMEOUT                   As Long = &H102
Public Const STATUS_PENDING                   As Long = &H103
Public Const STATUS_SEGMENT_NOTIFICATION      As Long = &H40000005
Public Const STATUS_GUARD_PAGE_VIOLATION      As Long = &H80000001
Public Const STATUS_DATATYPE_MISALIGNMENT     As Long = &H80000002
Public Const STATUS_BREAKPOINT                As Long = &H80000003
Public Const STATUS_SINGLE_STEP               As Long = &H80000004
Public Const STATUS_ACCESS_VIOLATION          As Long = &HC0000005
Public Const STATUS_IN_PAGE_ERROR             As Long = &HC0000006
Public Const STATUS_INVALID_HANDLE            As Long = &HC0000008
Public Const STATUS_NO_MEMORY                 As Long = &HC0000017
Public Const STATUS_ILLEGAL_INSTRUCTION       As Long = &HC000001D
Public Const STATUS_NONCONTINUABLE_EXCEPTION  As Long = &HC0000025
Public Const STATUS_INVALID_DISPOSITION       As Long = &HC0000026
Public Const STATUS_ARRAY_BOUNDS_EXCEEDED     As Long = &HC000008C
Public Const STATUS_FLOAT_DENORMAL_OPERAND    As Long = &HC000008D
Public Const STATUS_FLOAT_DIVIDE_BY_ZERO      As Long = &HC000008E
Public Const STATUS_FLOAT_INEXACT_RESULT      As Long = &HC000008F
Public Const STATUS_FLOAT_INVALID_OPERATION   As Long = &HC0000090
Public Const STATUS_FLOAT_OVERFLOW            As Long = &HC0000091
Public Const STATUS_FLOAT_STACK_CHECK         As Long = &HC0000092
Public Const STATUS_FLOAT_UNDERFLOW           As Long = &HC0000093
Public Const STATUS_INTEGER_DIVIDE_BY_ZERO    As Long = &HC0000094
Public Const STATUS_INTEGER_OVERFLOW          As Long = &HC0000095
Public Const STATUS_PRIVILEGED_INSTRUCTION    As Long = &HC0000096
Public Const STATUS_STACK_OVERFLOW            As Long = &HC00000FD
Public Const STATUS_CONTROL_C_EXIT            As Long = &HC000013A
Public Const STATUS_FLOAT_MULTIPLE_FAULTS     As Long = &HC00002B4
Public Const STATUS_FLOAT_MULTIPLE_TRAPS      As Long = &HC00002B5
Public Const STATUS_ILLEGAL_VLM_REFERENCE     As Long = &HC00002C0
