VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "constantAPI"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@PROJECT_LICENSE
Option Explicit
Option Base 0
Const CLASSID As String = "STDCONSTS"

Public Enum ExecutionState
    esSUCCESSFULL = 0
    esSUCCESS = SUCCESSFULL
    esFAILURE = 1
End Enum

Public Enum MathematicalConstants
    mPI = 3.14159
    mRADIAN = 1.74532
    mGRAD = 0.0157
End Enum

Public Enum Registry_HiveKeys
    rHKEY_CLASSES_ROOT = &H80000000
    rHKEY_CURRENT_USER = &H80000001
    rHKEY_LOCAL_MACHINE = &H80000002
    rHKEY_USERS = &H80000003
    rHKEY_PERFORMANCE_DATA = &H80000004
    rHKEY_CURRENT_CONFIG = &H80000005
    rHKEY_DYN_DATA = &H80000006
End Enum

Public Enum VarSize
    MAXLONG = 2147483647
    MINLONG = -2147483648#
    LONG_MAX = MAXLONG
    LONG_MIN = MINLONG

    MAXINTEGER = 32767
    MININTEGER = -32768
    INTEGER_MAX = MAXINTEGER
    INTEGER_MIN = MININTEGER
End Enum

Public Enum TimeDurations
    INFINITE = &HFFFFFFFF  ' Infinite timeout
    TIMEOUTINFINITE = &HFFFFFFFF  ' Infinite timeout
    INFINITETIMEOUT = &HFFFFFFFF  ' Infinite timeout
    MILLISECONDS = (1000)             ' 10 ^ 3
    NANOSECONDS = (1000000000)         ' 10 ^ 9
    UNITS = (10000000)        ' 10 ^ 7
End Enum

Public Enum TrueFalse
    YES = True
    NO = False
End Enum

Public Enum ItemState
    ITEMNOEXISTS = -1
    DEFAULTVALUE = -1
End Enum

Public Type HRESULT
    Value As Long
End Type
Public Type LRESULT
    Value As Long
End Type
Public Type wParam
    Value As Long
End Type
Public Type lParam
    Value As Long
End Type
Public Type CALLBACK
    Value As Long
End Type
Public Type HANDLE
    Value As Long
End Type
Public Type Ptr
    Value As Long
End Type
Public Type FILEHANDLE
    h As Long
    Path As String
End Type

Public Type DDWORD 'INT256
    HHigherValue1 As Long
    HLowerValue1 As Long
    LHigherValue1 As Long
    LLowerValue1 As Long
    HHigherValue2 As Long
    HLowerValue2 As Long
    LHigherValue2 As Long
    LLowerValue2 As Long
End Type
Public Type Int256
    HHigherValue1 As Long
    HLowerValue1 As Long
    LHigherValue1 As Long
    LLowerValue1 As Long
    HHigherValue2 As Long
    HLowerValue2 As Long
    LHigherValue2 As Long
    LLowerValue2 As Long
End Type
Public Type DQWORD 'INT128
    HHigherValue As Long
    HLowerValue As Long
    LHigherValue As Long
    LLowerValue As Long
End Type
Public Type Int128
    HHigherValue As Long
    HLowerValue As Long
    LHigherValue As Long
    LLowerValue As Long
End Type
Public Type QWORD 'INT64
    HigherValue As Long
    LowerValue As Long
End Type
Public Type Int64
    HigherValue As Long
    LowerValue As Long
End Type
Public Type DWORD 'INT32
    Value As Long
End Type
Public Type Int32
    Value As Long
End Type
Public Type WORD 'INT16
    Value As Integer
End Type
Public Type Int16
    Value As Integer
End Type
Public Type BYTE_ 'INT8
    Value As Byte
End Type
Public Type INT8
    Value As Byte
End Type

Dim bit_Mask(0 To 31) As Long

Private Sub Class_Initialize()
    Dim i As Long
    For i = 0 To 30
        bit_Mask(i) = 2 ^ i
    Next 'creating event mask
    bit_Mask(31) = &H80000000
End Sub

Public Function bitMask() As Long()
    bitMask = bit_Mask
End Function

Public Property Get HRESULT(v As HRESULT) As Long: HRESULT = v.Value: End Property
Public Property Let HRESULT(v As HRESULT, Value As Long):    v.Value = Value: End Property
Public Property Get LRESULT(v As LRESULT) As Long:    LRESULT = v.Value: End Property
Public Property Let LRESULT(v As LRESULT, Value As Long):    v.Value = Value: End Property
Public Property Get wParam(v As wParam) As Long:    wParam = v.Value: End Property
Public Property Let wParam(v As wParam, Value As Long):    v.Value = Value: End Property
Public Property Get lParam(v As lParam) As Long:    lParam = v.Value: End Property
Public Property Let lParam(v As lParam, Value As Long):    v.Value = Value: End Property
Public Property Get CALLBACK(v As CALLBACK) As Long:    CALLBACK = v.Value: End Property
Public Property Let CALLBACK(v As CALLBACK, Value As Long):    v.Value = Value: End Property
Public Property Get HANDLE(v As HANDLE) As Long:    HANDLE = v.Value: End Property
Public Property Let HANDLE(v As HANDLE, Value As Long):    v.Value = Value: End Property
Public Property Get Ptr(v As Ptr) As Long:    Ptr = v.Value: End Property
Public Property Let Ptr(v As Ptr, Value As Long):    v.Value = Value: End Property
