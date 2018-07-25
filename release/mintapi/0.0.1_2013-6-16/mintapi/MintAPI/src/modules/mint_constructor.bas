Attribute VB_Name = "mint_constructor"
Option Explicit



Public Function EventArgs(targetObject As Object) As EventArgs
    Dim E As New EventArgs
    Call E.Initialize(targetObject)
    Set EventArgs = E
End Function
Public Function DisposingEventArgs(targetObject As Object, Optional Continue As Boolean = True) As DisposingEventArgs
    Dim E As New DisposingEventArgs
    Call E.Initialize(targetObject, Continue)
    Set DisposingEventArgs = E
End Function
Public Function ExceptionOccuredEventArgs(targetObject As Object, Optional Exception As Exception) As ExceptionOccuredEventArgs
    Dim E As New ExceptionOccuredEventArgs
    Call E.Initialize(targetObject, Exception)
    Set ExceptionOccuredEventArgs = E
End Function
Public Function ApplicationMessageEventArgs(targetObject As Object, Args() As Variant) As ApplicationMessageEventArgs
    Dim E As New ApplicationMessageEventArgs
    Call E.Initialize(targetObject, Args)
    Set ApplicationMessageEventArgs = E
End Function
Public Function MString(str As String, Optional ByVal IsSpecialArgument As Boolean = False) As StringParser
    Dim sp As New StringParser
    Call sp.Initialize(str, IsSpecialArgument)
    Set MString = sp
End Function
Public Function StringParser(str As String, Optional ByVal IsSpecialArgument As Boolean = False) As StringParser
    Dim sp As New StringParser
    Call sp.Initialize(str, IsSpecialArgument)
    Set StringParser = sp
End Function

Public Function Class(Optional objClass As Object) As fclass
    Dim fc As New fclass
    Call fc.Initialize(objClass)
    Set Class = fc
End Function
Public Function fclass(Optional objClass As Object) As fclass
    Dim fc As New fclass
    Call fc.Initialize(objClass)
    Set fclass = fc
End Function
Public Function ByteArray(Optional target) As ByteArray
    Dim ba As New ByteArray
    Call ba.Initialize(target)
    Set ByteArray = ba
End Function

Public Function File(Optional Path As String = "") As File
    Dim f As New File
    Call f.Initialize(Path)
    Set File = f
End Function
Public Function Directory(Optional Path As String = "") As Directory
    Dim d As New Directory
    Call d.Initialize(Path)
    Set Directory = d
End Function
Public Function CurrentDirectory() As Directory
    Dim d As New Directory
    Set CurrentDirectory = d.CurrentDirectory
End Function
Public Function Registry(Optional Key As String = "") As Registry
    Dim R As New Registry
    Call R.Initialize(Key)
    Set Registry = R
End Function
Public Function Thread(Optional ByVal targetFunction As Method = Nothing) As Thread
    Dim t As New Thread
    Call t.Initialize(targetFuncHandle:=targetFunction)
    Set Thread = t
End Function
Public Function Method(ByVal Name As String, ByVal targetFunctionAddress As Long) As Method
    Dim M As New Method
    Call M.Initialize(Name, targetFunctionAddress)
    Set Method = M
End Function
Public Function CurrentThread() As Thread
    Dim t As New Thread
    Call t.Initialize
    Set CurrentThread = t
End Function
Public Function Process(Path As String, Arguments As String, Optional AsUser As String = "", Optional Environment As String) As Process
    Dim p As Process
    If AsUser = "" Then
        Set p = CurrentProcess.OpenProcess(Path, Arguments, Environment)
    Else
        Set p = CurrentProcess.OpenProcessAs(Path, CStr(AsUser), Arguments, Environment)
    End If
    Set Process = p
End Function
Public Function CurrentProcess() As Process
    Dim p As New Process
    Call p.Initialize
    Set CurrentProcess = p
End Function
Public Function Language(Optional Path As String = "") As Language
    Dim l As New Language
    Call l.Initialize(Path)
    Set Language = l
End Function
Public Function Configuration() As Configuration
    Dim Conf As New Configuration
    Call Conf.Initialize
    Set Configuration = Conf
End Function
Public Function MintAPIRegistry() As Registry
    Set MintAPIRegistry = Registry(APP_REGISTRYPATH)
End Function
Public Function BigNumber(Optional InitialValue) As BigNumber
    Dim bn As New BigNumber
    Call bn.Initialize(InitialValue)
    Set BigNumber = bn
End Function
Public Function BigNum(Optional InitialValue) As BigNumber
    Dim bn As New BigNumber
    Call bn.Initialize(InitialValue)
    Set BigNum = bn
End Function
Public Function NetAPI(Optional Arguments) As NetAPI
    Dim nAPI As New NetAPI
    Call nAPI.Initialize(Arguments)
    Set NetAPI = nAPI
End Function
Public Function Socket(Optional AddressFamily As AddressFamily = afUnspecified, Optional SocketType As SocketType = stStream, Optional Protocol As Protocol = pUnspecified) As Socket
    Dim Sock As New Socket
    Call Sock.Initialize(AddressFamily, SocketType, Protocol)
    Set Socket = Sock
End Function
Public Function GraphicObject(Optional Arguments) As GraphicMethods
    Dim GM As New GraphicMethods
    Call GM.Initialize(Arguments)
    Set GraphicObject = GM
End Function
Public Function GraphicMethods(Optional Arguments) As GraphicMethods
    Dim GM As New GraphicMethods
    Call GM.Initialize(Arguments)
    Set GraphicMethods = GM
End Function
Public Function GM(Optional Arguments) As GraphicMethods
    Dim GMr As New GraphicMethods
    Call GMr.Initialize(Arguments)
    Set GM = GMr
End Function

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
    If VarType(Value) = VBObject Then
        Set Argument.Value = Value
    Else
            Argument.Value = Value
    End If
End Function
Public Function CreateObjectBuffer(Name As String, ParamArray Args() As Variant) As Object
    Dim Arg() As Variant
    Arg = Args
    Dim OB As New ObjectBuffer
    Call OB.InitializeW(Name, Arg)
    Set CreateObjectBuffer = OB
End Function
Public Function CreateObjectBufferC(Name As String, Args() As Variant) As Object
    Dim OB As New ObjectBuffer
    Call OB.InitializeW(Name, Args)
    Set CreateObjectBufferC = OB
End Function
