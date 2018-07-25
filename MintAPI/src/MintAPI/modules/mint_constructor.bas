Attribute VB_Name = "mint_constructor"
Option Explicit

'=============================================
'=============================================
'=============================================
'<section Static Class Constructors>
'

'Public Function File() As File
'    Set File = Static_File
'End Function
'Public Function Method() As MethodStatic
'    Set Method = Static_MethodStatic
'End Function
''Public Function MintAPIRegistry() As Registry
''    Set MintAPIRegistry = Registry(APP_REGISTRYPATH)
''End Function
'Public Function Environment() As Environment
'    Set Environment = mint_static_mngr.Environment
'End Function
'Public Function Memory() As MemoryStatic
'    Set Memory = mint_static_mngr.Memory
'End Function
'Public Function Application() As Application
'    Set Application = mint_static_mngr.Application
'End Function
'Public Function Instance() As MintAPIInstance
'    Set Instance = mint_static_mngr.MintAPIInstance
'End Function
'Public Function Process() As ProcessStatic
'    Set Process = mint_static_mngr.ProcessStatic
'End Function
'Public Function Thread() As ThreadStatic
'    Set Thread = mint_static_mngr.ThreadStatic
'End Function
'Public Function Arrays() As ArrayStatic
'    Set Arrays = mint_static_mngr.ArrayStatic
'End Function
'Public Function Debugger() As Debugger
'    Set Debugger = mint_static_mngr.Debugger
'End Function
'Public Function Objects() As ObjectStatic
'    Set Objects = mint_static_mngr.ObjectStatic
'End Function
'Public Function MetaObject() As MetaObjectStatic
'    Set MetaObject = mint_static_mngr.MetaObjectStatic
'End Function
'Public Function Info() As Information
'    Set Info = mint_static_mngr.Information
'End Function
'Public Function Interlocked() As Interlocked
'    Set Interlocked = mint_static_mngr.Interlocked
'End Function
''Public Function Signals() As SignalStatic
''    Set Signals = mint_static_mngr.SignalsStatic
''End Function
'Public Function Signal() As SignalStatic
'    Set Signal = mint_static_mngr.SignalsStatic
'End Function
''Public Function Slot() As SlotStatic
''    Set Slot = mint_static_mngr.Slot
''End Function
'Public Function Configuration() As Configuration
'    Set Configuration = mint_static_mngr.Configuration
'End Function
'Public Function Provider() As ProviderStatic
'    Set Provider = mint_static_mngr.Provider
'End Function
'Public Function Runtime() As Runtime
'    Set Runtime = mint_static_mngr.Runtime
'End Function
'Public Function Console() As Console
'    Set Console = mint_static_mngr.Console
'End Function
'Public Function Version() As VersionStatic
'    Set Version = mint_static_mngr.VersionStatic
'End Function
'Public Function DBNull() As DBNullStatic
'    Set DBNull = mint_static_mngr.DBNullStatic
'End Function
'Public Function Exceptions() As Exceptions
'    Set Exceptions = mint_static_mngr.Exceptions
'End Function
'Public Function Enumerator() As EnumeratorStatic
'    Set Enumerator = mint_static_mngr.EnumeratorStatic
'End Function
'
'Public Function MathExt() As MathExt
'    Set MathExt = mint_static_mngr.MathExt
'End Function
'Public Function mString() As mString
'    Set mString = mint_static_mngr.mString
'End Function
'
'Public Function Path() As Path
'    Set Path = mint_static_mngr.Path
'End Function
'Public Function Directory() As Directory
'    Set Directory = mint_static_mngr.Directory
'End Function
'
'Public Function Registry() As Registry
'    Set Registry = mint_static_mngr.Registry
'End Function
'
'Public Function Buffer() As BufferStatic
'    Set Buffer = mint_static_mngr.BufferStatic
'End Function
'
'Public Function Guid() As GuidStatic
'    Set Guid = mint_static_mngr.GuidStatic
'End Function
'
'
'
'Public Function Comparer() As ComparerStatic
'    Set Comparer = mint_static_mngr.ComparerStatic
'End Function
'Public Function Default_Comparer() As IComparer
'    Set Default_Comparer = mint_static_mngr.Comparer
'End Function
'
'Public Function Convert() As Convert
'    Set Convert = mint_static_mngr.Convert
'End Function
'
'Public Function Culture() As CultureStatic
'    Set Culture = mint_static_mngr.Culture
'End Function
'Public Function TypeInfo() As TypeInfoStatic
'    Set TypeInfo = mint_static_mngr.TypeInfo
'End Function
'
'Public Function ThreadStack() As ThreadStackStatic
'    Set ThreadStack = mint_static_mngr.ThreadStack
'End Function
'
'Public Function Text() As TextStatic
'    Set Text = mint_static_mngr.Text
'End Function
'
'Public Function Timers() As TimerStatic
'    Set Timers = mint_static_mngr.Timer
'End Function
'
'Public Function Library() As LibraryStatic
'    Set Library = mint_static_mngr.Library
'End Function
'
'Public Function DateTime() As DateTimeStatic
'    Set DateTime = mint_static_mngr.DateTime
'End Function
'
'Public Function DynamicObject() As DynamicObjectStatic
'    Set DynamicObject = mint_static_mngr.DynamicObject
'End Function
'
'Public Function SharedMemory() As SharedMemoryStatic
'    Set SharedMemory = mint_static_mngr.SharedMemory
'End Function
'
'Public Function DiskDrive() As DiskDriveStatic
'    Set DiskDrive = mint_static_mngr.DiskDrive
'End Function
'
'Public Function Defaults() As Defaults
'    Set Defaults = mint_static_mngr.Defaults
'End Function
'
'Public Function Enumerable() As Enumerable
'    Set Enumerable = mint_static_mngr.Enumerable
'End Function
'
'Public Function Prototype() As MethodPrototypeStatic
'    Set Prototype = mint_static_mngr.Prototype
'End Function
'
'Public Function VirtualMemory() As VirtualMemoryStatic
'    Set VirtualMemory = mint_static_mngr.VirtualMemoryStatic
'End Function



'
'</section>
'---------------------------------------------
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Instanciate Class Constructors>
'
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
'Public Function CreateObjectBuffer(Name As String, ParamArray Args() As Variant) As Object
'    Dim Arg() As Variant
'    Arg = Args
'    Dim OB As New ObjectBuffer
'    Call OB.InitializeW(Name, Arg)
'    Set CreateObjectBuffer = OB
'End Function
'Public Function CreateObjectBufferC(Name As String, Args() As Variant) As Object
'    Dim OB As New ObjectBuffer
'    Call OB.InitializeW(Name, Args)
'    Set CreateObjectBufferC = OB
'End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section EventArgs Constructors>
'
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
'
'</section>
'---------------------------------------------
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section UDT Constructors>
'
Public Function Position(Left As Long, Top As Long) As Position
    Position.Left = Left
    Position.Top = Top
End Function
Public Function Size(Width As Long, Height As Long) As Size
    Size.Width = Width
    Size.Height = Height
End Function
Public Function Rectangle(Left As Long, Top As Long, Width As Long, Height As Long) As Rectangle
    Rectangle.Left = Left
    Rectangle.Top = Top
    Rectangle.Width = Width
    Rectangle.Height = Height
End Function
Public Function Point(X As Long, Y As Long) As Point
    Point.X = X
    Point.Y = Y
End Function
Public Function Point3D(X As Long, Y As Long, Z As Long) As Point3D
    Point3D.X = X
    Point3D.Y = Y
    Point3D.Z = Z
End Function
Public Function Region(Left As Long, Top As Long, Right As Long, Bottom As Long) As Region
    Region.Left = Left
    Region.Top = Top
    Region.Right = Right
    Region.Bottom = Bottom
End Function
Public Function Padding(Left As Long, Top As Long, Right As Long, Bottom As Long) As Padding
    Padding.Left = Left
    Padding.Top = Top
    Padding.Right = Right
    Padding.Bottom = Bottom
End Function
Public Function Margin(Left As Long, Top As Long, Right As Long, Bottom As Long) As Margin
    Margin.Left = Left
    Margin.Top = Top
    Margin.Right = Right
    Margin.Bottom = Bottom
End Function
Public Function Argument(Name As String, Value As Variant) As Argument
    Argument.Name = Name
    If IsObject(Value) Then
        Set Argument.Value = Value
    Else
            Argument.Value = Value
    End If
End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------
'---------------------------------------------

