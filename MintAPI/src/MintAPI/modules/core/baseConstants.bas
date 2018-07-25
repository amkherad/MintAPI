Attribute VB_Name = "baseConstants"
'@PROJECT_LICENSE

Option Explicit
Option Base 0

Public Const MAX_PATH = 260
Public Const MAX_PATH_A = MAX_PATH
Public Const MAX_PATH_W = 32767
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

Public Const MINTAPI_PATH As Long = MAX_PATH_W

Public Const CHAR_LEN As Long = 2

Public Const INVALID_HANDLE_VALUE As Long = -1

Public Const SUCCESSFULL As Long = 0
Public Const SUCCESS As Long = SUCCESSFULL
Public Const FAILURE As Long = -1
Public Const S_OK = &H0
Public Const vbFalse = &H0
Public Const vbTrue = &HFFFFFFFF

Public Const E_NOINTERFACE As Long = &H80004002

Public Const vbNullPtr As Long = 0

Public Const C_PI As Double = 3.14159265358979
'3.14159265358979323846264338327950288419716939937510582097494459230781640628620899862803482534211706798214808651328230664709384460955058223172535940812848111745028410270193852110555964462294895493038196442881097566593344612847564823378678316527120190914564856692346034861045432664821339360726024914127372458700660631558817488152092096282925409171536436789259036001133053054882046652138414695194151160943305727036575959195309218611738193261179310511854807446237996274956735188575272489122793818301194912983367336244065664308602139494639522473719070217986094370277053921717629317675238467481846766940513200056812714526356082778577134275778960917363717872146844090122495343014654958537105079227968925892354201995611212902196086403441815981362977477130996051870721134999999837297804995105973173281609631859502445945534690830264252230825334468503526193118817101000313783875288658753320838142061717766914730359825349042875546873115956286388235378759375195778185778053217122680661300192787661119590921642019893809525720106548586
Public Const C_RADIAN As Double = 1.74532925199433E-02
Public Const C_GRAD As Double = 0.015707963267949
Public Const E_Number As Double = 2.71828182845905 '2.71828182845904523536028747135266249775724709369995

Public Const MAXQWORD = ((9.22337203685477E+22) + 5807)   '9223372036854775807
Public Const MINQWORD = -((9.22337203685477E+22) + 5808)  '-9223372036854775808
Public Const QWORD_MAX = MAXQWORD
Public Const QWORD_MIN = MINQWORD

Public Const CB_KILOBYTES As Long = 1024
Public Const CB_MEGABYTES As Long = CB_KILOBYTES * 1024
Public Const CB_GIGABYTES As Long = CB_MEGABYTES * 1024
'Public Const CB_TERABYTES As Long = CB_GIGABYTES * 1024
'Public Const CB_PETABYTES As Long = CB_TERABYTES * 1024
'Public Const CB_EXABYTES  As Long = CB_PETABYTES * 1024
'Public Const CB_ZETTABYTES  As Long = CB_EXABYTES * 1024
'Public Const CB_YOTTABYTES  As Long = CB_ZETTABYTES * 1024

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

Public Const VLEN_PTR As Long = 4
Public Const VLEN_STRPTR As Long = 4
Public Const VLEN_BYTE As Long = 1
Public Const VLEN_INTEGER As Long = 2
Public Const VLEN_LONG As Long = 4
Public Const VLEN_SINGLE As Long = 4
Public Const VLEN_DOUBLE As Long = 8
Public Const VLEN_BOOLEAN As Long = 2
Public Const VLEN_CURRENCY As Long = 8
Public Const VLEN_DATE As Long = 8
Public Const VLEN_OBJECT As Long = 4
Public Const VLEN_UDTPTR As Long = 4
Public Const VLEN_VARIANT As Long = 16
Public Const VLEN_DECIMAL As Long = 16
Public Const VLEN_VARTYPE As Long = 2 'in variants.

Public Const FADF_AUTO          As Long = &H1
Public Const FADF_STATIC        As Long = &H2
Public Const FADF_EMBEDDED      As Long = &H4
Public Const FADF_FIXEDSIZE     As Long = &H10
Public Const FADF_RECORD        As Long = &H20
Public Const FADF_HAVEIID       As Long = &H40
Public Const FADF_HAVEVARTYPE   As Long = &H80
Public Const FADF_BSTR          As Long = &H100
Public Const FADF_UNKNOWN       As Long = &H200
Public Const FADF_DISPATCH      As Long = &H400
Public Const FADF_VARIANT       As Long = &H800
Public Const FADF_RESERVED      As Long = &HF008

Public Const C_MINTAPI_INTERNAL_ERROR As Long = 500

Public Const SYSTEM_DOTNET_KEY_ADDRESS As String = "HKEY_LOCALMACHINE\SOFTWARE\Microsoft\NET Framework Setup\NDP\"

Public Const C_INFINITE = &HFFFFFFFF  ' Infinite timeout
Public Const C_TIMEOUTINFINITE = &HFFFFFFFF  ' Infinite timeout
Public Const C_INFINITETIMEOUT = &HFFFFFFFF  ' Infinite timeout

Public Const C_MILLISECONDS = (1000)             ' 10 ^ 3
Public Const C_NANOSECONDS = (1000000000)         ' 10 ^ 9
Public Const C_UNITS = (C_NANOSECONDS / 100)        ' 10 ^ 7

Public Const C_ITEMNOTEXISTS As Long = -1
Public Const C_DEFAULTVALUE As Long = -1

Public Const MASK_BYTE0 As Long = &HFF
Public Const MASK_BYTE1 As Long = &HFF00
Public Const MASK_BYTE2 As Long = &HFF0000
Public Const MASK_BYTE3 As Long = &HFF000000
Public Const MASK_SHORT0 As Long = &HFFFF
Public Const MASK_SHORT1 As Long = &HFFFF0000
'
'Public Const ERROR_PATH_NOT_FOUND       As Long = 3
'Public Const ERROR_ACCESS_DENIED        As Long = 5
'Public Const ERROR_FILE_NOT_FOUND       As Long = 2
'Public Const ERROR_FILE_EXISTS          As Long = 80
'Public Const ERROR_INSUFFICIENT_BUFFER  As Long = 122

Public Const FILE_BEGIN = 0
Public Const FILE_CURRENT = 1
Public Const FILE_END = 2
Public Const FILE_EXECUTE = (&H20)

Public Const LOCALE_USER_DEFAULT As Long = 0

Public Const MEMBERID_NIL As Long = -1

Public Const sNewLine As String = vbNewLine


Public Const FUNC_ORDER_QueryInterface       As Long = 0
Public Const FUNC_ORDER_AddRef               As Long = 1
Public Const FUNC_ORDER_Release              As Long = 2

Public Const FUNC_ORDER_GetTypeInfoCount     As Long = 3
Public Const FUNC_ORDER_GetTypeInfo          As Long = 4
Public Const FUNC_ORDER_GetIDsOfNames        As Long = 5
Public Const FUNC_ORDER_Invoke               As Long = 6

Public Const FUNCNAME_IUNKNOWN_QueryInterface       As String = "QueryInterface"
Public Const FUNCNAME_IUNKNOWN_AddRef               As String = "AddRef"
Public Const FUNCNAME_IUNKNOWN_Release              As String = "Release"

Public Const FUNCNAME_IDISPATCH_GetTypeInfoCount     As String = "GetTypeInfoCount"
Public Const FUNCNAME_IDISPATCH_GetTypeInfo          As String = "GetTypeInfo"
Public Const FUNCNAME_IDISPATCH_GetIDsOfNames        As String = "GetIDsOfNames"
Public Const FUNCNAME_IDISPATCH_Invoke               As String = "Invoke"

Public Const FUNCNAME_IEnumVariant_NewEnum       As String = "NewEnum"
Public Const FUNCNAME_IEnumerable_GetEnumerator       As String = "GetEnumerator"

'========================================
'   Com Interface GUIDs
'========================================
' Basic Interfaces
Public Const GUIDS_NULL                      As Long = 0
Public Const GUIDS_IUnknown                  As String = "00000000-0000-0000-C000-000000000046"
Public Const GUIDS_IDispatch                 As String = "00020400-0000-0000-C000-000000000046"
Public Const GUIDS_ITypeInfo                 As String = "00020401-0000-0000-C000-000000000046"
Public Const GUIDS_ITypeInfo2                As String = "00020412-0000-0000-C000-000000000046"
Public Const GUIDS_IRecordInfo               As String = "0000002F-0000-0000-C000-000000000046"
Public Const GUIDS_ITypeMARSHAL              As String = "0000002D-0000-0000-C000-000000000046"
Public Const GUIDS_ITypeLib                  As String = "00020402-0000-0000-C000-000000000046"
Public Const GUIDS_ITypeLib2                 As String = "00020411-0000-0000-C000-000000000046"
Public Const GUIDS_ITypeFactory              As String = "0000002E-0000-0000-C000-000000000046"
Public Const GUIDS_ITypeComp                 As String = "00020403-0000-0000-C000-000000000046"
Public Const GUIDS_ITypeChangeEvents         As String = "00020410-0000-0000-C000-000000000046"
Public Const GUIDS_ICreateTypeInfo           As String = "00020405-0000-0000-C000-000000000046"
Public Const GUIDS_ICreateTypeInfo2          As String = "0002040E-0000-0000-C000-000000000046"
Public Const GUIDS_ICreateTypeLib            As String = "00020406-0000-0000-C000-000000000046"
Public Const GUIDS_ICreateTypeLib2           As String = "0002040F-0000-0000-C000-000000000046"
Public Const GUIDS_IOleObject                As String = "00000112-0000-0000-C000-000000000046"
Public Const GUIDS_IContinue                 As String = "0000012a-0000-0000-C000-000000000046"
Public Const GUIDS_IDropSource               As String = "00000121-0000-0000-C000-000000000046"
Public Const GUIDS_IDropTarget               As String = "00000122-0000-0000-C000-000000000046"
Public Const GUIDS_IEnumOLEVERB              As String = "00000104-0000-0000-C000-000000000046"
Public Const GUIDS_IOleAdviseHolder          As String = "00000111-0000-0000-C000-000000000046"
Public Const GUIDS_IOleCache                 As String = "0000011e-0000-0000-C000-000000000046"
Public Const GUIDS_IOleCache2                As String = "00000128-0000-0000-C000-000000000046"
Public Const GUIDS_IOleCacheControl          As String = "00000129-0000-0000-C000-000000000046"
Public Const GUIDS_IOleClientSite            As String = "00000118-0000-0000-C000-000000000046"
Public Const GUIDS_IOleContainer             As String = "0000011b-0000-0000-C000-000000000046"
Public Const GUIDS_IOleInPlaceActiveObject   As String = "00000117-0000-0000-C000-000000000046"
Public Const GUIDS_IOleInPlaceFrame          As String = "00000116-0000-0000-C000-000000000046"
Public Const GUIDS_IOleInPlaceObject         As String = "00000113-0000-0000-C000-000000000046"
Public Const GUIDS_IOleInPlaceSite           As String = "00000119-0000-0000-C000-000000000046"
Public Const GUIDS_IOleInPlaceUIWindow       As String = "00000115-0000-0000-C000-000000000046"
Public Const GUIDS_IOleItemContainer         As String = "0000011c-0000-0000-C000-000000000046"
Public Const GUIDS_IOleLink                  As String = "0000011d-0000-0000-C000-000000000046"
Public Const GUIDS_IOleWindow                As String = "00000114-0000-0000-C000-000000000046"
Public Const GUIDS_IParseDisplayName         As String = "0000011a-0000-0000-C000-000000000046"
' Ole
Public Const GUIDS_IViewObject               As String = "0000010d-0000-0000-C000-000000000046"
Public Const GUIDS_IViewObject2              As String = "00000127-0000-0000-C000-000000000046"
Public Const GUIDS_IEnumVARIANT              As String = "00020404-0000-0000-C000-000000000046"
Public Const GUIDS_IAddrExclusionControl     As String = "00000148-0000-0000-C000-000000000046"
Public Const GUIDS_IAddrTrackingControl      As String = "00000147-0000-0000-C000-000000000046"
Public Const GUIDS_IAdviseSink               As String = "0000010f-0000-0000-C000-000000000046"
Public Const GUIDS_IAdviseSink2              As String = "00000125-0000-0000-C000-000000000046"
Public Const GUIDS_IAsyncManager             As String = "0000002A-0000-0000-C000-000000000046"
Public Const GUIDS_ICancelMethodCalls        As String = "00000029-0000-0000-C000-000000000046"
Public Const GUIDS_IComThreadingInfo         As String = "000001ce-0000-0000-C000-000000000046"
Public Const GUIDS_IDataAdviseHolder         As String = "00000110-0000-0000-C000-000000000046"
Public Const GUIDS_IDataObject               As String = "0000010e-0000-0000-C000-000000000046"
Public Const GUIDS_IEnumFORMATETC            As String = "00000103-0000-0000-C000-000000000046"
Public Const GUIDS_IEnumMoniker              As String = "00000102-0000-0000-C000-000000000046"
Public Const GUIDS_IEnumSTATDATA             As String = "00000105-0000-0000-C000-000000000046"
Public Const GUIDS_IEnumSTATSTG              As String = "0000000d-0000-0000-C000-000000000046"
Public Const GUIDS_IEnumString               As String = "00000101-0000-0000-C000-000000000046"
Public Const GUIDS_IEnumUnknown              As String = "00000100-0000-0000-C000-000000000046"
Public Const GUIDS_IExternalConnection       As String = "00000019-0000-0000-C000-000000000046"
Public Const GUIDS_IBindCtx                  As String = "0000000E-0000-0000-C000-000000000046"
Public Const GUIDS_IStdMarshalInfo           As String = "00000018-0000-0000-C000-000000000046"
Public Const GUIDS_IStorage                  As String = "0000000b-0000-0000-C000-000000000046"
Public Const GUIDS_IStream                   As String = "0000000c-0000-0000-C000-000000000046"
Public Const GUIDS_ISurrogate                As String = "00000022-0000-0000-C000-000000000046"
Public Const GUIDS_ISurrogateService         As String = "000001d4-0000-0000-C000-000000000046"
Public Const GUIDS_ISynchronize              As String = "00000030-0000-0000-C000-000000000046"
Public Const GUIDS_ISynchronizeContainer     As String = "00000033-0000-0000-C000-000000000046"
Public Const GUIDS_ISynchronizeEvent         As String = "00000032-0000-0000-C000-000000000046"
Public Const GUIDS_ISynchronizeHandle        As String = "00000031-0000-0000-C000-000000000046"
Public Const GUIDS_ISynchronizeMutex         As String = "00000025-0000-0000-C000-000000000046"
Public Const GUIDS_IUrlMon                   As String = "00000026-0000-0000-C000-000000000046"
Public Const GUIDS_IWaitMultiple             As String = "0000002B-0000-0000-C000-000000000046"
Public Const GUIDS_AsyncIAdviseSink          As String = "00000150-0000-0000-C000-000000000046"
Public Const GUIDS_AsyncIAdviseSink2         As String = "00000151-0000-0000-C000-000000000046"
Public Const GUIDS_AsyncIMultiQI             As String = "000E0020-0000-0000-C000-000000000046"
' Tool Interfaces
Public Const GUIDS_ISupportErrorInfo         As String = "DF0B3D60-548F-101B-8E65-08002B2BD119"
Public Const GUIDS_IPropertyBag              As String = "55272A00-42CB-11CE-8135-00AA004BB851"
Public Const GUIDS_IErrorLog                 As String = "3127CA40-446E-11CE-8135-00AA004BB851"
Public Const GUIDS_IErrorInfo                As String = "1CF2B120-547D-101B-8E65-08002B2BD119"
Public Const GUIDS_ICreateErrorInfo          As String = "22F03340-547D-101B-8E65-08002B2BD119"
Public Const GUIDS_IAsyncRpcChannelBuffer    As String = "a5029fb6-3c34-11d1-9c99-00c04fb998aa"
Public Const GUIDS_IBlockingLock             As String = "30f3d47a-6447-11d1-8e3c-00c04fb9386d"
Public Const GUIDS_ICallFactory              As String = "1c733a30-2a1c-11ce-ade5-00aa0044773d"
Public Const GUIDS_IDirectWriterLock         As String = "0e6d4d92-6738-11cf-9608-00aa00680db4"
Public Const GUIDS_IDummyHICONIncluder       As String = "947990de-cc28-11d2-a0f7-00805f858fb1"
Public Const GUIDS_IFillLockBytes            As String = "99caf010-415e-11cf-8814-00aa00b569f5"
Public Const GUIDS_ILayoutStorage            As String = "0e6d4d90-6738-11cf-9608-00aa00680db4"
Public Const GUIDS_IOplockStorage            As String = "8d19c834-8879-11d1-83e9-00c04fc2c6d4"
Public Const GUIDS_IPipeByte                 As String = "DB2F3ACA-2F86-11d1-8E04-00C04FB9989A"
Public Const GUIDS_IPipeDouble               As String = "DB2F3ACE-2F86-11d1-8E04-00C04FB9989A"
Public Const GUIDS_IPipeLong                 As String = "DB2F3ACC-2F86-11d1-8E04-00C04FB9989A"
Public Const GUIDS_IProcessInitControl       As String = "72380d55-8d2b-43a3-8513-2b6ef31434e9"
Public Const GUIDS_IProgressNotify           As String = "a9d758a0-4617-11cf-95fc-00aa00680db4"
Public Const GUIDS_IPSFactoryBuffer          As String = "D5F569D0-593B-101A-B569-08002B2DBF7A"
Public Const GUIDS_IReleaseMarshalBuffers    As String = "eb0cb9e8-7996-11d2-872e-0000f8080859"
Public Const GUIDS_IROTData                  As String = "f29f6bc0-5021-11ce-aa15-00006901293f"
Public Const GUIDS_IRpcChannelBuffer         As String = "D5F56B60-593B-101A-B569-08002B2DBF7A"
Public Const GUIDS_IRpcChannelBuffer2        As String = "594f31d0-7f19-11d0-b194-00a0c90dc8bf"
Public Const GUIDS_IRpcChannelBuffer3        As String = "25B15600-0115-11d0-BF0D-00AA00B8DFD2"
Public Const GUIDS_IRpcProxyBuffer           As String = "D5F56A34-593B-101A-B569-08002B2DBF7A"
Public Const GUIDS_IRpcStubBuffer            As String = "D5F56AFC-593B-101A-B569-08002B2DBF7A"
Public Const GUIDS_IRpcSyntaxNegotiate       As String = "58a08519-24c8-4935-b482-3fd823333a4f"
Public Const GUIDS_ISequentialStream         As String = "0c733a30-2a1c-11ce-ade5-00aa0044773d"
Public Const GUIDS_IThumbnailExtractor       As String = "969dc708-5c76-11d1-8d86-0000f804b057"
Public Const GUIDS_ITimeAndNoticeControl     As String = "bc0bf6ae-8878-11d1-83e9-00c04fc2c6d4"
Public Const GUIDS_AsyncIPipeByte            As String = "DB2F3ACB-2F86-11d1-8E04-00C04FB9989A"
Public Const GUIDS_AsyncIPipeDouble          As String = "DB2F3ACF-2F86-11d1-8E04-00C04FB9989A"
Public Const GUIDS_AsyncIPipeLong            As String = "DB2F3ACD-2F86-11d1-8E04-00C04FB9989A"


' Ascii table by Kelly Ethridge VBCorLib:modConstants.bas #2004
Public Const vbUpperA           As Long = &H41
Public Const vbLowerA           As Long = &H61
Public Const vbLowerD           As Long = &H64
Public Const vbUpperD           As Long = &H44
Public Const vbLowerF           As Long = &H66
Public Const vbUpperF           As Long = &H46
Public Const vbLowerG           As Long = &H67
Public Const vbUpperG           As Long = &H47
Public Const vbLowerH           As Long = &H68
Public Const vbUpperH           As Long = &H48
Public Const vbLowerM           As Long = &H6D
Public Const vbUpperM           As Long = &H4D
Public Const vbLowerR           As Long = &H72
Public Const vbUpperR           As Long = &H52
Public Const vbLowerS           As Long = &H73
Public Const vbLowerT           As Long = &H74
Public Const vbUpperT           As Long = &H54
Public Const vbLowerU           As Long = &H75
Public Const vbUpperU           As Long = &H55
Public Const vbLowerY           As Long = &H79
Public Const vbUpperY           As Long = &H59
Public Const vbUpperZ           As Long = &HFA
Public Const vbLowerZ           As Long = &H7A
Public Const vbZero             As Long = &H30
Public Const vbOne              As Long = &H31
Public Const vbFive             As Long = &H35
Public Const vbNine             As Long = &H39
Public Const vbPlus             As Long = &H2B
Public Const vbMinus            As Long = &H2D
Public Const vbBackSlash        As Long = &H5C
Public Const vbForwardSlash     As Long = &H2F
Public Const vbColon            As Long = &H3A
Public Const vbSemiColon        As Long = &H3B
Public Const vbEqual            As Long = &H3D
Public Const vbReturn           As Long = &HD
Public Const vbLineFeed         As Long = &HA
Public Const vbSpace            As Long = &H20
Public Const vbPound            As Long = &H23
Public Const vbDollar           As Long = &H24
Public Const vbPercent          As Long = &H25
Public Const vbDoubleQuote      As Long = &H22
Public Const vbSingleQuote      As Long = &H27
Public Const vbComma            As Long = &H2C
Public Const vbPeriod           As Long = &H2E
Public Const vbInvalidChar      As Long = &HFFFFFFFF
' String versions
Public Const vbColonS           As String = ":"
Public Const vbSemiColonS       As String = ";"
Public Const vbBackSlashS       As String = "\"
Public Const vbForwardSlashS    As String = "/"
Public Const vbPeriodS          As String = "."


Public Const SYSTEMEXCEPTION_NONCONTINUABLE      As Long = &H1    ' Noncontinuable exception
Public Const SYSTEMEXCEPTION_MAXIMUM_PARAMETERS  As Long = 15     ' maximum number of exception parameters


'These are obsoletes.
'Public Type HRESULT
'    Value As Long
'End Type
'Public Type LRESULT
'    Value As Long
'End Type
'Public Type wParam
'    Value As Long
'End Type
'Public Type lParam
'    Value As Long
'End Type
'Public Type Callback
'    Value As Long
'End Type
'Public Type API_SECURITY_ATTRIBUTES
'    nLength As Long
'    lpSecurityDescriptor As Long
'    bInheritHandle As Long
'End Type

'Public Type DDWORD 'INT256
'    LLowerValue2 As Long
'    LHigherValue2 As Long
'    HLowerValue2 As Long
'    HHigherValue2 As Long
'    LLowerValue1 As Long
'    LHigherValue1 As Long
'    HLowerValue1 As Long
'    HHigherValue1 As Long
'End Type
'Public Type Int256
'    LLowerValue2 As Long
'    LHigherValue2 As Long
'    HLowerValue2 As Long
'    HHigherValue2 As Long
'    LLowerValue1 As Long
'    LHigherValue1 As Long
'    HLowerValue1 As Long
'    HHigherValue1 As Long
'End Type
'Public Type DQWORD 'INT128
'    LLowerValue As Long
'    LHigherValue As Long
'    HLowerValue As Long
'    HHigherValue As Long
'End Type
'Public Type Int128
'    LLowerValue As Long
'    LHigherValue As Long
'    HLowerValue As Long
'    HHigherValue As Long
'End Type
'Public Type QWORD 'INT64
'    LowerValue As Long
'    HigherValue As Long
'End Type
'Public Type Int64
'    HigherValue As Long
'    LowerValue As Long
'End Type
'Public Type DWORD 'INT32
'    Value As Long
'End Type
'Public Type Int32
'    Value As Long
'End Type
'Public Type WORD 'INT16
'    Value As Integer
'End Type
'Public Type Int16
'    Value As Integer
'End Type

'These are obsoletes.

'Dim inited As Boolean
'Public bitMask(0 To 31) As Long
'
'Public Sub Initialize()
'    If inited Then Exit Sub
'    Dim i As Long
'    For i = 0 To 30
'        bitMask(i) = 2 ^ i
'    Next 'creating event mask
'    bitMask(31) = &H80000000
'    inited = True
'End Sub
'Public Sub Dispose(Optional ByVal Force As Boolean = False)
'    If Not inited Then Exit Sub
'    Erase bitMask
'    inited = False
'End Sub
