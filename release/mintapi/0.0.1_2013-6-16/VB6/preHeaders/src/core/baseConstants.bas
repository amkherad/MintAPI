Attribute VB_Name = "baseConstants"
'@PROJECT_LICENSE

Option Explicit
Option Base 0
Const CLASSID As String = "baseConstants"

'COMPILERCONDITION = 0
'COMPILERCONDITION_WIN32VERSION
'COMPILERCONDITION_WIN32SPECIALPATH
'COMPILERCONDITION_FILEDIRECTORIES
'COMPILERCONDITION_BINARIES
'COMPILERCONDITION_STRINGS

Public Const MAX_PATH = 260
Public Const MAXNLEN As Long = 256
Public Const TINYLPSTR As Long = 64
Public Const TINY2LPSTR As Long = 96
Public Const MINILPSTR As Long = 128
Public Const MINI2LPSTR As Long = 192
Public Const SMALLLPSTR As Long = 256
Public Const LARGELPSTR As Long = 512
Public Const LARGE2LPSTR As Long = 768
Public Const XLARGELPSTR As Long = 1024
Public Const MULTILPSTR As Long = 2048

Public Const SUCCESSFULL As Long = 0
Public Const SUCCESS As Long = SUCCESSFULL

Public Const PI As Double = 3.14159265358979
'3.14159265358979323846264338327950288419716939937510582097494459230781640628620899862803482534211706798214808651328230664709384460955058223172535940812848111745028410270193852110555964462294895493038196442881097566593344612847564823378678316527120190914564856692346034861045432664821339360726024914127372458700660631558817488152092096282925409171536436789259036001133053054882046652138414695194151160943305727036575959195309218611738193261179310511854807446237996274956735188575272489122793818301194912983367336244065664308602139494639522473719070217986094370277053921717629317675238467481846766940513200056812714526356082778577134275778960917363717872146844090122495343014654958537105079227968925892354201995611212902196086403441815981362977477130996051870721134999999837297804995105973173281609631859502445945534690830264252230825334468503526193118817101000313783875288658753320838142061717766914730359825349042875546873115956286388235378759375195778185778053217122680661300192787661119590921642019893809525720106548586
Public Const RADIAN As Double = 1.74532925199433E-02
Public Const GRAD As Double = 0.015707963267949

Public Const MAXQWORD = ((9.22337203685477E+22) + 5807)   '9223372036854775807
Public Const MINQWORD = -((9.22337203685477E+22) + 5808)  '-9223372036854775808
Public Const QWORD_MAX = MAXQWORD
Public Const QWORD_MIN = MINQWORD

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006

Public Const MAXLONG As Long = &H7FFFFFFF
Public Const MINLONG As Long = &H80000000
Public Const LONG_MAX As Long = MAXLONG
Public Const LONG_MIN As Long = MINLONG

Public Const MAXINTEGER As Integer = 32767
Public Const MININTEGER As Integer = -32768
Public Const INTEGER_MAX As Integer = MAXINTEGER
Public Const INTEGER_MIN As Integer = MININTEGER

Public Const DEFAULTERROR As Long = 500
Public Const DEFAULT_ERROR As Long = 500

Public Const SYSTEM_DOTNET_KEY_ADDRESS As String = "HKEY_LOCALMACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\"
Public Const SYSTEM_RoNOTE_ADDRESS As String = "HKEY_LOCALMACHINE\SOFTWARE\RoNOTE"
Public Const SYSTEM_yScript_ADDRESS As String = "HKEY_LOCALMACHINE\SOFTWARE\yScript"
Public Const SYSTEM_geLib_ADDRESS As String = "HKEY_LOCALMACHINE\SOFTWARE\geLib"

Public Const INFINITE = &HFFFFFFFF  ' Infinite timeout
Public Const TIMEOUTINFINITE = &HFFFFFFFF  ' Infinite timeout
Public Const INFINITETIMEOUT = &HFFFFFFFF  ' Infinite timeout

Public Const iTRUE As Byte = True
Public Const iFALSE As Byte = False
Public Const YES As Byte = True
Public Const NO As Byte = False

Public Const ERRORS_NOTSUPPORTEDOS As String = "OS Version Not Supported."
Public Const ERRORS_ACCESSDENIED As String = "Access Denied."
Public Const ERRORS_UNKNOWNVALUE As String = "Unknown Argument Value."
Public Const ERRORS_ARGUMENTLENGTHOVERLOAD As String = "Argument Length Overloaded."
Public Const ERRORS_STRINGLENGTHOVERLOADED As String = "String Length Is Too Much."
Public Const ERRORS_INVALIDCALL As String = "Invalid Call Method Or Some Preset Properties Are Invalid Or Some Argument Are Invalid."
Public Const ERRORS_SYSTEMCALLFAILURE As String = "An Error Occured When Trying To Call A System Method."
Public Const ERRORS_ZERONUMBER As String = "Argument Must Not Equal To Zero."
Public Const ERRORS_NEGATIVEZERONUMBER As String = "Argument Must Be Grater Than Zero."
Public Const ERRORS_NEGATIVENUMBER As String = "Argument Must Be Grater Or Equal To Zero."
Public Const ERRORS_INVALIDARGUMENTVALUE As String = "Argument Must Not Equal To Zero."
Public Const ERRORS_INVALIDARGUMENT As String = "Invalid Argument Type."
Public Const ERRORS_TOOMANYARGUMENTS As String = "Too Many Arguments Passed."
Public Const ERRORS_AFEWARGUMENTS As String = "A Few Arguments Passed."
Public Const ERRORS_INVALIDARGUMENTTYPE As String = "Invalid Argument Type."
Public Const ERRORS_ARGUMENTNULL As String = "Argument Is Null."
Public Const ERRORS_NOTIMPLEMENTED As String = "Not Implemented."
Public Const ERRORS_INVALIDLBOUND As String = "Invalid Array Lower Bound."
Public Const ERRORS_INVALIDUBOUND As String = "Invalid Array Upper Bound."
Public Const ERRORS_VALUETOOHIGH As String = "Value Is Too High."
Public Const ERRORS_VALUETOOLOW As String = "Value Is Too Low."
Public Const ERRORS_CLASSNOTINITIALIZED As String = "Class Does Not Initialized."
Public Const ERRORS_ITEMDOESNOTEXISTS As String = "Item Does Not Exists."
Public Const ERRORS_ITEMEXISTS As String = "Item Exists."
Public Const ERRORS_LISTISEMPTY As String = "List Is Empty."
Public Const ERRORS_BUFFERISEMPTY As String = "Buffer Is Empty."
Public Const ERRORS_ERROR As String = "An Error Occured."
Public Const ERRORS_CANCELEDERROR As String = "Operation Canceled."
Public Const ERRORS_DISPOSED As String = "Control Have Been Disposed."
Public Const ERRORS_OUTOFRANGE As String = "Subscript Out Of Range."
Public Const ERRORS_INVALIDNAME As String = "Invalid Name."
Public Const ERRORS_INVALIDVAR As String = "Invalid Variable."
Public Const ERRORS_INVALIDVARTYPE As String = "Invalid Variable Type."
Public Const ERRORS_INVALIDADDRESS As String = "Invalid Address."
Public Const ERRORS_INVALIDPATH As String = "Invalid Path."
Public Const ERRORS_INVALIDFILE As String = "Invalid File Type."
Public Const ERRORS_FILENOTFOUND As String = "File Not Found."
Public Const ERRORS_FILEEXISTS As String = "File Already Exists."
Public Const ERRORS_FILENOTEXISTS As String = "File Does Not Exists."
Public Const ERRORS_PATHNOTEXISTS As String = "Path Does Not Exists."
Public Const ERRORS_CANTOPENFILE As String = "Error In Opening File."
Public Const ERRORS_CANTCLOSEFILE As String = "Error In Closing File."
Public Const ERRORS_CANTREADFILE As String = "Error Reading From File."
Public Const ERRORS_CANTWRITEFILE As String = "Error Writing In File."
Public Const ERRORS_INVALIDSTATUS As String = "Invalid Target Status."
Public Const ERRORS_INVALIDHANDLE As String = "Invalid Handle."
Public Const ERRORS_TARGETNOTREADY As String = "Target Not Ready."
Public Const ERRORS_TARGETNOTOPENED As String = "Target Not Opened."
Public Const ERRORS_SYSTEMCOMPOSITONNOTENABLED As String = "System Composition Not Enabled."

Public Const Milliseconds = (1000)             ' 10 ^ 3
Public Const NANOSECONDS = (1000000000)         ' 10 ^ 9
Public Const UNITS = (NANOSECONDS / 100)        ' 10 ^ 7

Public Const ITEMNOEXISTS As Long = -1
Public Const DefaultValue As Long = -1

Public Const MASK_BYTE0 As Long = &HFF
Public Const MASK_BYTE1 As Long = &HFF00
Public Const MASK_BYTE2 As Long = &HFF0000
Public Const MASK_BYTE3 As Long = &HFF000000
Public Const MASK_SHORT0 As Long = &HFFFF
Public Const MASK_SHORT1 As Long = &HFFFF0000

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
Public Type CallBack
    Value As Long
End Type
Public Type Handle
    Value As Long
End Type
Public Type Ptr
    Value As Long
End Type
Public Type FileHandle
    h As Long
    Path As String
End Type
Public Type API_SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Long
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

Dim inited As Boolean
Public bitMask(0 To 31) As Long

Public Sub Initialize()
    If inited Then Exit Sub
    Dim i As Long
    For i = 0 To 30
        bitMask(i) = 2 ^ i
    Next 'creating event mask
    bitMask(31) = &H80000000
    inited = True
End Sub
Public Sub Dispose(Optional ByVal Force As Boolean = False)
    If Not inited Then Exit Sub
    Erase bitMask
    inited = False
End Sub


Public Function APISecurityAttributes(lpSecurityDescriptor As Long, bInheritHandle As Long) As API_SECURITY_ATTRIBUTES
    APISecurityAttributes.bInheritHandle = bInheritHandle
    APISecurityAttributes.lpSecurityDescriptor = lpSecurityDescriptor
    APISecurityAttributes.nLength = Len(APISecurityAttributes)
End Function

Public Property Get HRESULT(v As HRESULT) As Long: HRESULT = v.Value: End Property
Public Property Let HRESULT(v As HRESULT, Value As Long):    v.Value = Value: End Property
Public Property Get LRESULT(v As LRESULT) As Long:    LRESULT = v.Value: End Property
Public Property Let LRESULT(v As LRESULT, Value As Long):    v.Value = Value: End Property
Public Property Get wParam(v As wParam) As Long:    wParam = v.Value: End Property
Public Property Let wParam(v As wParam, Value As Long):    v.Value = Value: End Property
Public Property Get lParam(v As lParam) As Long:    lParam = v.Value: End Property
Public Property Let lParam(v As lParam, Value As Long):    v.Value = Value: End Property
Public Property Get CallBack(v As CallBack) As Long:    CallBack = v.Value: End Property
Public Property Let CallBack(v As CallBack, Value As Long):    v.Value = Value: End Property
Public Property Get Handle(v As Handle) As Long:    Handle = v.Value: End Property
Public Property Let Handle(v As Handle, Value As Long):    v.Value = Value: End Property
Public Property Get Ptr(v As Ptr) As Long:    Ptr = v.Value: End Property
Public Property Let Ptr(v As Ptr, Value As Long):    v.Value = Value: End Property
