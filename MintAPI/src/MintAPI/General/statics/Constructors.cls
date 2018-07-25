VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Constructors"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"ApplicationMessageEventArgs"
Attribute VB_Ext_KEY = "Member1" ,"ArgumentList"
Attribute VB_Ext_KEY = "Member2" ,"ByteArray"
Attribute VB_Ext_KEY = "Member3" ,"File"
Attribute VB_Ext_KEY = "Member4" ,"List"
Attribute VB_Ext_KEY = "Member5" ,"StringParser"
Attribute VB_Ext_KEY = "Member6" ,"Timer"
Attribute VB_Ext_KEY = "Member7" ,"Vector"
Attribute VB_Ext_KEY = "Member8" ,"BigNumber"
Attribute VB_Ext_KEY = "Member9" ,"Library"
Attribute VB_Ext_KEY = "Member10" ,"ThreadStack"
Attribute VB_Ext_KEY = "Member11" ,"TypeInfo"
Attribute VB_Ext_KEY = "Member12" ,"Culture"
Attribute VB_Ext_KEY = "Member13" ,"Comparer"
Attribute VB_Ext_KEY = "Member14" ,"Guid"
Attribute VB_Ext_KEY = "Member15" ,"Registry"
Attribute VB_Ext_KEY = "Member16" ,"Buffer"
Attribute VB_Ext_KEY = "Member17" ,"Convert"
Attribute VB_Ext_KEY = "Member18" ,"mString"
Attribute VB_Ext_KEY = "Member19" ,"MathExt"
Attribute VB_Ext_KEY = "Member20" ,"DBNull"
Attribute VB_Ext_KEY = "Member21" ,"Version"
Attribute VB_Ext_KEY = "Member22" ,"Console"
Attribute VB_Ext_KEY = "Member23" ,"Runtime"
Attribute VB_Ext_KEY = "Member24" ,"Provider"
Attribute VB_Ext_KEY = "Member25" ,"Configuration"
Attribute VB_Ext_KEY = "Member26" ,"Exceptions"
Attribute VB_Ext_KEY = "Member27" ,"Stack"
Attribute VB_Ext_KEY = "Member28" ,"Language"
Attribute VB_Ext_KEY = "Member29" ,"Thread"
Attribute VB_Ext_KEY = "Member30" ,"Debugger"
Attribute VB_Ext_KEY = "Member31" ,"Directory"
Attribute VB_Ext_KEY = "Member32" ,"Path"
Attribute VB_Ext_KEY = "Member33" ,"Process"
Attribute VB_Ext_KEY = "Member34" ,"OperatingSystem"
Attribute VB_Ext_KEY = "Member35" ,"Generics"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member36" ,"Environment"
Attribute VB_Ext_KEY = "Member37" ,"Application"
Attribute VB_Ext_KEY = "Member38" ,"Memory"
Attribute VB_Ext_KEY = "Member39" ,"Method"
Attribute VB_Ext_KEY = "Member40" ,"ExceptionOccuredEventArgs"
Attribute VB_Ext_KEY = "Member41" ,"CancelEventArgs"
Attribute VB_Ext_KEY = "Member42" ,"EventArgs"
Attribute VB_Ext_KEY = "Member43" ,"Argument"
Option Explicit

'==============================================================================
'== Type Constructors Section =================================================
'==============================================================================
Public Function Int64(ByVal HigherValue As Long, ByVal LowerValue As Long) As Int64
    Int64.HigherValue = HigherValue
    Int64.LowerValue = LowerValue
End Function
Public Function Position(ByVal Left As Long, ByVal Top As Long) As Position
    Position.Left = Left
    Position.Top = Top
End Function
Public Function Size(ByVal Width As Long, ByVal Height As Long) As Size
    Size.Width = Width
    Size.Height = Height
End Function
Public Function Rectangle(ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long) As Rectangle
    Rectangle.Left = Left
    Rectangle.Top = Top
    Rectangle.Width = Width
    Rectangle.Height = Height
End Function
Public Function Point(ByVal X As Long, ByVal Y As Long) As Point
    Point.X = X
    Point.Y = Y
End Function
Public Function Point3D(ByVal X As Long, ByVal Y As Long, ByVal Z As Long) As Point3D
    Point3D.X = X
    Point3D.Y = Y
    Point3D.Z = Z
End Function
Public Function Region(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long) As Region
    Region.Left = Left
    Region.Top = Top
    Region.Right = Right
    Region.Bottom = Bottom
End Function
Public Function Padding(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long) As Padding
    Padding.Left = Left
    Padding.Top = Top
    Padding.Right = Right
    Padding.Bottom = Bottom
End Function
Public Function Margin(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long) As Margin
    Margin.Left = Left
    Margin.Top = Top
    Margin.Right = Right
    Margin.Bottom = Bottom
End Function
Public Function Argument(ByVal Name As String, ByRef Value As Variant) As Argument
    Set Argument = New Argument
    Call Argument.Constructor0(Name, Value)
End Function
'Public Function IPAddress(IPAddressValue As String) As IPAddress
'    IPAddress.IPAddress = IPAddressValue
'End Function
'Public Function SocketAddress(IPAddress As String, SocketPort As Long) As SocketAddress
'    SocketAddress.IPAddress = IPAddress
'    SocketAddress.SocketPort = SocketPort
'End Function
'Public Function EndPoint(IPAddress As String, SocketPort As Long, Optional AddressFamily As AddressFamily = afUnspecified) As EndPoint
'    EndPoint.IPAddress = IPAddress
'    EndPoint.SocketPort = SocketPort
'    EndPoint.Tag = 0
'    EndPoint.AddressFamily = AddressFamily
'End Function

'==============================================================================
'== EventArgs Constructors Section ============================================
'==============================================================================
Public Function EventArgs(ByVal Sender As Object) As EventArgs
    Set EventArgs = New EventArgs
    Call EventArgs.Constructor0(Sender)
End Function
Public Function CancelEventArgs(ByVal Sender As Object, ByVal Cancel As Boolean) As CancelEventArgs
    Set CancelEventArgs = New CancelEventArgs
    Call CancelEventArgs.Constructor0(Sender, Cancel)
End Function
Public Function ExceptionOccuredEventArgs(ByVal Sender As Object, ByVal Exception As Exception) As ExceptionOccuredEventArgs
    Set ExceptionOccuredEventArgs = New ExceptionOccuredEventArgs
    Call ExceptionOccuredEventArgs.Constructor0(Sender, Exception)
End Function
'Public Function ApplicationMessageEventArgs(ByVal Sender As Object, Args() As Variant) As ApplicationMessageEventArgs
'    Set ApplicationMessageEventArgs = New ApplicationMessageEventArgs
'    Call ApplicationMessageEventArgs.Constructor0(Sender, True)
'End Function

'==============================================================================
'== Exceptions Constructors Section ===========================================
'==============================================================================

'Public Function Operation(Key As String) As Operation
'    Set Operation = mod_application.Operation(Key)
'End Function


'==============================================================================
'== Static Classes Constructors Section =======================================
'==============================================================================

Public Function File() As File
    Set File = mint_static_mngr.File
End Function
Public Function Method() As MethodStatic
    Set Method = mint_static_mngr.Method
End Function
'Public Function Delegate(ByVal Reference As Long, Optional ByVal CallingConvension As MethodCallingConvension) As Method
'    Set Delegate = Method.FromReference(Reference, CallingConvension)
'End Function
Public Function Objects() As ObjectStatic
    Set Objects = mint_static_mngr.Objects
End Function
Public Function AL() As ArgumentListStatic
    Set AL = mint_static_mngr.AL
End Function
Public Function Info() As Information
    Set Info = mint_static_mngr.Info
End Function
Public Function Memory() As MemoryStatic
    Set Memory = mint_static_mngr.Memory
End Function
Public Function Application() As Application
    Set Application = mint_static_mngr.Application
End Function
Public Function Environment() As Environment
    Set Environment = mint_static_mngr.Environment
End Function
Public Function Interlocked() As Interlocked
    Set Interlocked = mint_static_mngr.Interlocked
End Function
Public Function Generics() As Generics
    Set Generics = mint_static_mngr.Generics
End Function
Public Function OS() As OperatingSystem
    Set OS = mint_static_mngr.OS
End Function
Public Function Process() As ProcessStatic
    Set Process = mint_static_mngr.Process
End Function

Public Function Timers() As TimerStatic
    Set Timers = mint_static_mngr.Timer
End Function

Public Function Path() As Path
    Set Path = mint_static_mngr.Path
End Function
Public Function Directory() As Directory
    Set Directory = mint_static_mngr.Directory
End Function
'Public Function Enums() As Enums
'    Set Enums = mint_static_mngr.Enums
'End Function
Public Function Instance() As MintAPIInstance
    Set Instance = mint_static_mngr.Instance
End Function
'Public Function Graphics() As GraphicMethods
'    Set Graphics = mint_static_mngr.GraphicMethods
'End Function
Public Function Debugger() As Debugger
    Set Debugger = mint_static_mngr.Debugger
End Function
Public Function Arrays() As ArrayStatic
    Set Arrays = mint_static_mngr.Arrays
End Function
Public Function Thread() As ThreadStatic
    Set Thread = mint_static_mngr.Thread
End Function
'Public Function MetaObject() As MetaObjectStatic
'    Set MetaObject = mint_static_mngr.MetaObject
'End Function
Public Function Signal() As SignalStatic
    Set Signal = mint_static_mngr.Signal
End Function
Public Function Exps() As Exceptions
    Set Exps = mint_static_mngr.Exps
End Function
'Public Function Slot() As SlotStatic
'    Set Slot = mint_static_mngr.Slot
'End Function
Public Function Configuration() As Configuration
    Set Configuration = mint_static_mngr.Configuration
End Function
Public Function Provider() As ProviderStatic
    Set Provider = mint_static_mngr.Provider
End Function
Public Function Runtime() As Runtime
    Set Runtime = mint_static_mngr.Runtime
End Function
Public Function Console() As Console
    Set Console = mint_static_mngr.Console
End Function
Public Function Version() As VersionStatic
    Set Version = mint_static_mngr.Version
End Function
Public Function DBNull() As DBNullStatic
    Set DBNull = mint_static_mngr.DBNull
End Function
Public Function Enumerator() As EnumeratorStatic
    Set Enumerator = mint_static_mngr.Enumerator
End Function
Public Function GC() As GarbageCollector
    Set GC = mint_static_mngr.GC
End Function

Public Function mMath() As mMath
    Set mMath = mint_static_mngr.mMath
End Function
Public Function MathExt() As mMath
    Set MathExt = mint_static_mngr.mMath
End Function
Public Function mString() As mString
    Set mString = mint_static_mngr.mString
End Function
Public Function Text() As TextStatic
    Set Text = mint_static_mngr.Text
End Function

Public Function Convert() As Convert
    Set Convert = mint_static_mngr.Convert
End Function

Public Function Buffer() As BufferStatic
    Set Buffer = mint_static_mngr.Buffer
End Function

Public Function Registry() As Registry
    Set Registry = mint_static_mngr.Registry
End Function

Public Function Guid() As GuidStatic
    Set Guid = mint_static_mngr.Guid
End Function

Public Function Comparer() As ComparerStatic
    Set Comparer = mint_static_mngr.Comparer
End Function

Public Function Culture() As CultureStatic
    Set Culture = mint_static_mngr.Culture
End Function
Public Function TypeInfo() As TypeInfoStatic
    Set TypeInfo = mint_static_mngr.TypeInfo
End Function

Public Function ThreadStack() As ThreadStackStatic
    Set ThreadStack = mint_static_mngr.Thread
End Function

Public Function Library() As LibraryStatic
    Set Library = mint_static_mngr.Library
End Function

Public Function DateTime() As DateTimeStatic
    Set DateTime = mint_static_mngr.DateTime
End Function

Public Function DynamicObject() As DynamicObjectStatic
    Set DynamicObject = mint_static_mngr.DynamicObject
End Function

Public Function SharedMemory() As SharedMemoryStatic
    Set SharedMemory = mint_static_mngr.SharedMemory
End Function

Public Function DiskDrive() As DiskDriveStatic
    Set DiskDrive = mint_static_mngr.DiskDrive
End Function

Public Function Defaults() As Defaults
    Set Defaults = mint_static_mngr.Defaults
End Function

Public Function Enumerable() As Enumerable
    Set Enumerable = mint_static_mngr.Enumerable
End Function

Public Function Prototype() As MethodPrototypeStatic
    Set Prototype = mint_static_mngr.Prototype
End Function

Public Function VirtualMemory() As VirtualMemoryStatic
    Set VirtualMemory = mint_static_mngr.VirtualMemory
End Function

Public Function IUnknown() As IUnknownHelper
    Set IUnknown = mint_static_mngr.IUnknown
End Function

'==============================================================================
'== Class Constructors Section ================================================
'==============================================================================

Public Function BasicResult(ByRef ResultExpression As Variant) As IResult
    Dim Basic As New BasicResult
    Call Basic.Constructor0(ResultExpression)
    Set BasicResult = Basic
End Function

Public Function ByteArray(Optional Expression, Optional ByVal ConvertToBinary As Boolean = True, Optional Length As Long = -1) As ByteArray
    Dim BA As New ByteArray
    Call BA.Constructor1(Expression, ConvertToBinary, Length)
    Set ByteArray = BA
End Function
'Public Function BigNumber(Optional InitialValue) As BigNumber
''    Dim BN As New BigNumber
''    Call BN.Initialize(InitialValue)
''    Set BigNumber = BN
'End Function
'Public Function BigNum(Optional InitialValue) As BigNumber
''    Dim BN As New BigNumber
''    Call BN.Initialize(InitialValue)
''    Set BigNum = BN
'End Function

'==========================================================================================
'------------------------------------------------------------------------------------------
'..........................................................................................

