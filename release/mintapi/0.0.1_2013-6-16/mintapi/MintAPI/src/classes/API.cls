VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "API"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Global MintAPI methods provided for user applications."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Graphics"
Attribute VB_Ext_KEY = "Member1" ,"Pixmap"
Attribute VB_Ext_KEY = "Member2" ,"GraphicMethods"
Attribute VB_Ext_KEY = "Member3" ,"Console"
Attribute VB_Ext_KEY = "Member4" ,"ExAPI"
Attribute VB_Ext_KEY = "Member5" ,"Algorithms"
Attribute VB_Ext_KEY = "Member6" ,"Vector"
Attribute VB_Ext_KEY = "Member7" ,"List"
Attribute VB_Ext_KEY = "Member8" ,"ArgumentList"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member9" ,"Language"
Attribute VB_Ext_KEY = "Member10" ,"Configuration"
Attribute VB_Ext_KEY = "Member11" ,"Process"
Attribute VB_Ext_KEY = "Member12" ,"Method"
Attribute VB_Ext_KEY = "Member13" ,"Thread"
Attribute VB_Ext_KEY = "Member14" ,"Timer"
Attribute VB_Ext_KEY = "Member15" ,"Registry"
Attribute VB_Ext_KEY = "Member16" ,"File"
Attribute VB_Ext_KEY = "Member17" ,"Directory"
Attribute VB_Ext_KEY = "Member18" ,"StringParser"
Attribute VB_Ext_KEY = "Member19" ,"NetAPI"
Attribute VB_Ext_KEY = "Member20" ,"Socket"
Attribute VB_Ext_KEY = "Member21" ,"BigNumber"
Attribute VB_Ext_KEY = "Member22" ,"fclass"
Attribute VB_Ext_KEY = "Member23" ,"Memory"
Attribute VB_Ext_KEY = "Member24" ,"ByteArray"
Attribute VB_Ext_KEY = "Member25" ,"DiskDrive"
Attribute VB_Ext_KEY = "Member26" ,"Generics"
Attribute VB_Ext_KEY = "Member27" ,"Environment"
Attribute VB_Ext_KEY = "Member28" ,"SpecialMethods"
Attribute VB_Ext_KEY = "Member29" ,"Application"
Attribute VB_Ext_KEY = "Member30" ,"ApplicationMessageEventArgs"
Attribute VB_Ext_KEY = "Member31" ,"ExceptionOccuredEventArgs"
Attribute VB_Ext_KEY = "Member32" ,"DisposingEventArgs"
Attribute VB_Ext_KEY = "Member33" ,"EventArgs"
'@PROJECT_LICENSE

Option Base 0
Option Explicit
Const CLASSID As String = "API"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


Implements IAliasable


Dim dbg_filter_level As Long

'==========================================================================================
'== Global Methods Section ================================================================
'==========================================================================================
Public Sub InitializeMintAPIEnvironment()
    '
End Sub
Public Function tr(Key As String) As String
    tr = mod_application.CurrentLanguage.translate(Key)
End Function
Public Function GetAddressOf(Address As Long) As Long
    GetAddressOf = Address
End Function


Public Sub PutDataToByteArrayN(ByteArray() As Byte, Data, Optional Length As Long = -1)
    ByteArray = mint_get_byte_array_of(Data, Length)
End Sub
Public Sub PutDataToByteArray(outByteArray As ByteArray, Data, Optional Length As Long = -1)
    Set outByteArray = ByteArray(mint_get_byte_array_of(Data, Length))
End Sub
Public Sub PutByteArrayToDataN(Data, ByteArray() As Byte, Optional Length As Long = -1)
    Call mint_put_byte_array_to(Data, ByteArray, Length)
End Sub
Public Sub PutByteArrayToData(Data, ByteArray As ByteArray, Optional Length As Long = -1)
    Call mint_put_byte_array_to(Data, ByteArray.constData, Length)
End Sub


Public Function ArraySize(targetArray) As Long
If Not (VarType(targetArray) And vbArray) = vbArray Then throw InvalidArgumentTypeException
    On Error GoTo zeroLength
    ArraySize = (UBound(targetArray) - LBound(targetArray) + 1)
zeroLength:
End Function
Public Function IsEmptyArray(targetArray) As Boolean
If Not (VarType(targetArray) And vbArray) = vbArray Then throw InvalidArgumentTypeException
    On Error GoTo zeroLength
    IsEmptyArray = (UBound(targetArray) - LBound(targetArray) + 1) <= 0
    Exit Function
zeroLength:
    IsEmptyArray = True
End Function
Public Sub EmptyVar(Variable)
    If IsObject(Variable) Then
        Set Variable = Nothing
    Else
        If IsArray(Variable) Then
            Erase Variable
        Else
            Variable = Empty
        End If
    End If
End Sub
Public Function IsEmptyVariable(targetVariable) As Boolean
    Select Case VarType(targetVariable)
        Case vbArray
            IsEmptyVariable = (IsArrayEmpty(targetVariable))
        Case VBObject
            IsEmptyVariable = (targetVariable Is Nothing)
        Case Else
            IsEmptyVariable = (targetVariable = Empty)
    End Select
End Function

Public Function Array_String(ParamArray str() As Variant) As String()
    Dim i As Long, strSize As Long, retVal() As String
    On Error GoTo zeroLength
    strSize = UBound(str) - LBound(str) + 1
zeroLength:
    If strSize > 0 Then
        ReDim retVal(strSize - 1)
        For i = 0 To strSize - 1
            retVal(i) = CStr(str(i))
        Next
    End If
    Array_String = retVal
End Function
Public Function Array_Object(ParamArray Objects() As Variant) As Object()
    Dim i As Long, objSize As Long, retVal() As Object
    On Error GoTo zeroLength
    objSize = UBound(Objects) - LBound(Objects) + 1
zeroLength:
    If objSize > 0 Then
        ReDim retVal(objSize - 1)
        For i = 0 To objSize - 1
            Set retVal(i) = Objects(i)
        Next
    End If
    Array_Object = retVal
End Function
Public Function Array_Double(ParamArray Doubles() As Variant) As Double()
    Dim i As Long, dblSize As Long, retVal() As Double
    On Error GoTo zeroLength
    dblSize = UBound(Doubles) - LBound(Doubles) + 1
zeroLength:
    If dblSize > 0 Then
        ReDim retVal(dblSize - 1)
        For i = 0 To dblSize - 1
            retVal(i) = CDbl(Doubles(i))
        Next
    End If
    Array_Double = retVal
End Function
Public Function Array_Single(ParamArray Singles() As Variant) As Single()
    Dim i As Long, sngSize As Long, retVal() As Single
    On Error GoTo zeroLength
    sngSize = UBound(Singles) - LBound(Singles) + 1
zeroLength:
    If sngSize > 0 Then
        ReDim retVal(sngSize - 1)
        For i = 0 To sngSize - 1
            retVal(i) = CSng(Singles(i))
        Next
    End If
    Array_Single = retVal
End Function
Public Function Array_Long(ParamArray Longs() As Variant) As Long()
    Dim i As Long, lngSize As Long, retVal() As Long
    On Error GoTo zeroLength
    lngSize = UBound(Longs) - LBound(Longs) + 1
zeroLength:
    If lngSize > 0 Then
        ReDim retVal(lngSize - 1)
        For i = 0 To lngSize - 1
            retVal(i) = CLng(Longs(i))
        Next
    End If
    Array_Long = retVal
End Function
Public Function Array_Integer(ParamArray Ints() As Variant) As Integer()
    Dim i As Long, intSize As Long, retVal() As Integer
    On Error GoTo zeroLength
    intSize = UBound(Ints) - LBound(Ints) + 1
zeroLength:
    If intSize > 0 Then
        ReDim retVal(intSize - 1)
        For i = 0 To intSize - 1
            retVal(i) = CLng(Ints(i))
        Next
    End If
    Array_Integer = retVal
End Function
Public Function Array_Byte(ParamArray Bytes() As Variant) As Byte()
    Dim i As Long, btSize As Long, retVal() As Byte
    On Error GoTo zeroLength
    btSize = UBound(Bytes) - LBound(Bytes) + 1
zeroLength:
    If btSize > 0 Then
        ReDim retVal(btSize - 1)
        For i = 0 To btSize - 1
            retVal(i) = CByte(Bytes(i))
        Next
    End If
    Array_Byte = retVal
End Function
Public Function s(ParamArray str() As Variant) As String()
    Dim i As Long, strSize As Long, retVal() As String
    On Error GoTo zeroLength
    strSize = UBound(str) - LBound(str) + 1
zeroLength:
    If strSize > 0 Then
        ReDim retVal(strSize - 1)
        For i = 0 To strSize - 1
            retVal(i) = CStr(str(i))
        Next
    End If
    s = retVal
End Function
Public Function StringArray(ParamArray str() As Variant) As String()
    Dim i As Long, strSize As Long, retVal() As String
    On Error GoTo zeroLength
    strSize = UBound(str) - LBound(str) + 1
zeroLength:
    If strSize > 0 Then
        ReDim retVal(strSize - 1)
        For i = 0 To strSize - 1
            retVal(i) = CStr(str(i))
        Next
    End If
    StringArray = retVal
End Function
Public Function ByteArrayArray(ParamArray bas() As Variant) As ByteArray()
    Dim i As Long, baSize As Long, retVal() As ByteArray
    On Error GoTo zeroLength
    baSize = UBound(bas) - LBound(bas) + 1
zeroLength:
    On Error GoTo 0
    If baSize > 0 Then
        ReDim retVal(baSize - 1)
        For i = 0 To baSize - 1
            If VarType(bas(i)) <> VBObject Then throw InvalidArgumentTypeException("Only ByteArrays Accepted.")
            If Not TypeOf bas(i) Is ByteArray Then throw InvalidArgumentTypeException("Only ByteArrays Accepted.")
            Set retVal(i) = bas(i)
        Next
    End If
    ByteArrayArray = retVal
End Function
Public Function StringArrayToVariantArray(strs() As String) As Variant()
    Dim strsCount As Long, retVars() As Variant
    strsCount = ArraySize(strs)
    If strsCount > 0 Then
        ReDim retVars(strsCount - 1)
        Dim i As Long, xCounter As Long
        For i = LBound(strs) To UBound(strs)
            retVars(xCounter) = CStr(strs(i))
            xCounter = xCounter + 1
        Next
    End If
    StringArrayToVariantArray = retVars
End Function
'==========================================================================================
'------------------------------------------------------------------------------------------
'..........................................................................................

'==========================================================================================
'== Debugging Section =================================================================
'==========================================================================================
Public Sub FilterDebuggerMessages(Optional ByVal FilterLevel As Long = 0)
    If FilterLevel < 0 Then FilterLevel = 0
    dbg_filter_level = FilterLevel
End Sub
Public Sub Debugger(Message As String, Optional ByVal Level As Long = 1, Optional ByVal targetApplication As IApplication)
    If Level < 1 Then Level = 1
    Call Log(Message, vbRed, targetApplication)
    If Level <= dbg_filter_level Then Exit Sub

End Sub
Public Sub Log(Message As String, Optional ByVal Color As Long = vbBlack, Optional ByVal targetApplication As IApplication)
    If targetApplication Is Nothing Then Set targetApplication = tApplication
    If tApplication Is Nothing Then throw InvalidStatusException
    'Format(NOW,"MM\DD\YYYY hh:mm:ss")
End Sub
'==========================================================================================
'------------------------------------------------------------------------------------------
'..........................................................................................

'==========================================================================================
'== Streaming API Section =================================================================
'==========================================================================================
Public Sub out(ByVal Stream As ITargetStream, Data, Optional Length As Long = -1)
    If Stream Is Nothing Then Set Stream = appOutput
    If Not Stream.getState(sdOutStream) Then throw InvalidStatusException
    Call Stream.outStream(Data, Length)
End Sub
Public Sub inp(ByVal Stream As ITargetStream, Data, Optional Length As Long = -1)
    If Stream Is Nothing Then Set Stream = appInput
    If Not Stream.getState(sdInStream) Then throw InvalidStatusException
    Call Stream.inStream(Data, Length)
End Sub
Public Sub lout(ByVal Stream As ITargetStream, ParamArray Data() As Variant)
    If Stream Is Nothing Then Set Stream = appOutput
    If Not Stream.getState(sdOutStream) Then throw InvalidStatusException
    Dim Size As Long
    On Error GoTo zeroLength
    Size = UBound(Data) - LBound(Data) + 1
    If Size = 0 Then Exit Sub
    Dim i As Long
    On Error GoTo 0
    For i = LBound(Data) To UBound(Data)
        Call Stream.outStream(Data(i))
    Next
zeroLength:
End Sub
Public Sub linp(ByVal Stream As ITargetStream, ParamArray Data() As Variant)
    If Stream Is Nothing Then Set Stream = appInput
    If Not Stream.getState(sdInStream) Then throw InvalidStatusException
    Dim Size As Long
    On Error GoTo zeroLength
    Size = UBound(Data) - LBound(Data) + 1
    If Size = 0 Then Exit Sub
    Dim i As Long
    On Error GoTo 0
    For i = LBound(Data) To UBound(Data)
        Call Stream.inStream(Data(i))
    Next
zeroLength:
End Sub
Public Function endl() As Object
    Dim obj As New ObjectBuffer
    Call obj.Initialize(OBJECTBUFFER_STREAMINGVALUE, OBJECTBUFFER_VALUES_STREAMING_NEWLINE)
    Set endl = obj
End Function
Public Function sendl() As Object
    Dim obj As New ObjectBuffer
    Call obj.Initialize(OBJECTBUFFER_STREAMINGVALUE, OBJECTBUFFER_VALUES_STREAMING_NEWLINE)
    Set sendl = obj
End Function
Public Function sclr() As Object
    Dim obj As New ObjectBuffer
    Call obj.Initialize(OBJECTBUFFER_STREAMINGVALUE, OBJECTBUFFER_VALUES_STREAMING_CLEAR)
    Set sclr = obj
End Function
Public Function sflush() As Object
    Dim obj As New ObjectBuffer
    Call obj.Initialize(OBJECTBUFFER_STREAMINGVALUE, OBJECTBUFFER_VALUES_STREAMING_FLUSH)
    Set sflush = obj
End Function
Public Function sseek(Index As Long) As Object
    Dim obj As New ObjectBuffer
    Call obj.Initialize(OBJECTBUFFER_STREAMINGVALUE, OBJECTBUFFER_VALUES_STREAMING_SEEK, Index)
    Set sseek = obj
End Function

Public Sub stdin(ParamArray Data())
    Dim Size As Long
    On Error GoTo zeroLength
    Size = UBound(Data) - LBound(Data) + 1
zeroLength:
    If Size = 0 Then Exit Sub
    Dim i As Long
    On Error GoTo 0
    For i = LBound(Data) To UBound(Data)
        Call mod_application.stdin(Data(i))
    Next
End Sub
Public Sub stdout(ParamArray Data())
    Dim Size As Long
    On Error GoTo zeroLength
    Size = UBound(Data) - LBound(Data) + 1
zeroLength:
    If Size = 0 Then Exit Sub
    Dim i As Long
    On Error GoTo 0
    For i = LBound(Data) To UBound(Data)
        Call mod_application.stdout(Data(i))
    Next
End Sub
Public Sub stderr(ParamArray Data())
    Dim Size As Long
    On Error GoTo zeroLength
    Size = UBound(Data) - LBound(Data) + 1
zeroLength:
    If Size = 0 Then Exit Sub
    Dim i As Long
    On Error GoTo 0
    For i = LBound(Data) To UBound(Data)
        Call mod_application.stderr(Data(i))
    Next
End Sub
Public Sub stdWarning(ParamArray Data())
    Dim Size As Long
    On Error GoTo zeroLength
    Size = UBound(Data) - LBound(Data) + 1
zeroLength:
    If Size = 0 Then Exit Sub
    Dim i As Long
    On Error GoTo 0
    For i = LBound(Data) To UBound(Data)
        Call mod_application.stdWarning(Data(i))
    Next
End Sub
Public Sub stdDebug(ParamArray Data())
    Dim Size As Long
    On Error GoTo zeroLength
    Size = UBound(Data) - LBound(Data) + 1
zeroLength:
    If Size = 0 Then Exit Sub
    Dim i As Long
    On Error GoTo 0
    For i = LBound(Data) To UBound(Data)
        Call mod_application.stdDebug(Data(i))
    Next
End Sub
Public Sub stdError(ParamArray Data())
    Dim Size As Long
    On Error GoTo zeroLength
    Size = UBound(Data) - LBound(Data) + 1
zeroLength:
    If Size = 0 Then Exit Sub
    Dim i As Long
    On Error GoTo 0
    For i = LBound(Data) To UBound(Data)
        Call mod_application.stdError(Data(i))
    Next
End Sub
Public Sub stdCritical(ParamArray Data())
    Dim Size As Long
    On Error GoTo zeroLength
    Size = UBound(Data) - LBound(Data) + 1
zeroLength:
    If Size = 0 Then Exit Sub
    Dim i As Long
    On Error GoTo 0
    For i = LBound(Data) To UBound(Data)
        Call mod_application.stdCritical(Data(i))
    Next
End Sub

Public Sub conout(ParamArray Data())
    Dim Size As Long
    On Error GoTo zeroLength
    Size = UBound(Data) - LBound(Data) + 1
zeroLength:
    If Size = 0 Then Exit Sub
    Dim i As Long
    On Error GoTo 0
    Dim sConsole As ITargetStream
    Set sConsole = New Console
    Call sConsole.OpenStream(sdOutStream)
    For i = LBound(Data) To UBound(Data)
        Call mod_application.out(sConsole, Data(i))
    Next
End Sub
Public Sub conin(ParamArray Data())
    Dim Size As Long
    On Error GoTo zeroLength
    Size = UBound(Data) - LBound(Data) + 1
zeroLength:
    If Size = 0 Then Exit Sub
    Dim i As Long
    On Error GoTo 0
    Dim sConsole As ITargetStream
    Set sConsole = New Console
    Call sConsole.OpenStream(sdInStream)
    For i = LBound(Data) To UBound(Data)
        Call mod_application.out(sConsole, Data(i))
    Next
End Sub
'==========================================================================================
'------------------------------------------------------------------------------------------
'..........................................................................................

'==========================================================================================
'== Constructors Section ==================================================================
'==========================================================================================
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
Public Function IPAddress(IPAddressValue As String) As IPAddress
    IPAddress.IPAddress = IPAddressValue
End Function
Public Function SocketAddress(IPAddress As String, SocketPort As Long) As SocketAddress
    SocketAddress.IPAddress = IPAddress
    SocketAddress.SocketPort = SocketPort
End Function
Public Function EndPoint(IPAddress As String, SocketPort As Long, Optional AddressFamily As AddressFamily = afUnspecified) As EndPoint
    EndPoint.IPAddress = IPAddress
    EndPoint.SocketPort = SocketPort
    EndPoint.Tag = 0
    EndPoint.AddressFamily = AddressFamily
End Function
Public Function MIDIMessage(Volume As Long) As MIDIMessage
    MIDIMessage.Volume = Volume
End Function



Public Function EventArgs(targetObject As Object) As EventArgs
    Set EventArgs = mint_constructor.EventArgs(targetObject)
End Function
Public Function DisposingEventArgs(targetObject As Object, Optional Continue As Boolean = True) As DisposingEventArgs
    Set DisposingEventArgs = mint_constructor.DisposingEventArgs(targetObject, Continue)
End Function
Public Function ExceptionOccuredEventArgs(targetObject As Object, Optional Exception As Exception = Nothing) As DisposingEventArgs
    Set ExceptionOccuredEventArgs = mint_constructor.ExceptionOccuredEventArgs(targetObject, Exception)
End Function
Public Function ApplicationMessageEventArgs(targetObject As Object, Args() As Variant) As ApplicationMessageEventArgs
    Set ApplicationMessageEventArgs = mint_constructor.ApplicationMessageEventArgs(targetObject, Args)
End Function

'Public Function Operation(Key As String) As Operation
'    Set Operation = mod_application.Operation(Key)
'End Function
Public Function Exceptions(Optional Arguments) As Exceptions
    Dim excp As New Exceptions
    Call excp.Initialize(Arguments)
    Set Exceptions = excp
End Function

Public Function Application(Optional Arguments) As Application
    Dim appS As New Application
    Call appS.Initialize(Arguments)
    Set Application = appS
End Function
Public Function SpecialMethods() As SpecialMethods
    Dim sm As New SpecialMethods
    Call sm.Initialize
    Set SpecialMethods = sm
End Function
Public Function Environment(Optional Args) As Environment
    Dim E As New Environment
    Call E.Initialize(Args)
    Set Environment = E
End Function
Public Function Generics() As Generics
    Dim gn As New Generics
    Call gn.Initialize
    Set Generics = gn
End Function
Public Function Machine() As MachineControl
    Dim mc As New MachineControl
    Call mc.Initialize
    Set Machine = mc
End Function
Public Function OS() As MachineControl
    Dim mc As New MachineControl
    Call mc.Initialize
    Set OS = mc
End Function
Public Function DiskDrive(Optional DriveLetter) As DiskDrive
    Dim dd As New DiskDrive
    Call dd.Initialize(DriveLetter)
    Set DiskDrive = dd
End Function

Public Function ByteArray(Optional target) As ByteArray
    Dim BA As New ByteArray
    Call BA.Initialize(target)
    Set ByteArray = BA
End Function
Public Function Memory(Optional target) As Memory
    Dim M As New Memory
    Call M.Initialize(target)
    Set Memory = M
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
Public Function Socket(Optional AddressFamily As AddressFamily = afUnspecified, Optional SocketType As SocketType = stStream, Optional Protocol As Protocol = pUnspecified) As Socket
    Dim Sock As New Socket
    Call Sock.Initialize(AddressFamily, SocketType, Protocol)
    Set Socket = Sock
End Function
Public Function NetAPI(Optional Arguments) As NetAPI
    Dim nAPI As New NetAPI
    Call nAPI.Initialize(Arguments)
    Set NetAPI = nAPI
End Function

Public Function MString(Optional str As String, Optional ByVal IsSpecialArgument As Boolean = False) As StringParser
    Dim sp As New StringParser
    Call sp.Initialize(str, IsSpecialArgument)
    Set MString = sp
End Function
Public Function StringParser(Optional str As String, Optional ByVal IsSpecialArgument As Boolean = False) As StringParser
    Dim sp As New StringParser
    Call sp.Initialize(str, IsSpecialArgument)
    Set StringParser = sp
End Function
Public Function Parser(Optional Arg, Optional ByVal IsSpecialArgument As Boolean = False)
    If Not IsMissing(Arg) Then
        Select Case VarType(Arg)
            Case vbString, vbDecimal, vbByte, vbBoolean, vbInteger, vbLong, vbLongDate, vbLongTime, vbDate, vbDouble, vbSingle, vbCurrency

            Case Else
                throw InvalidArgumentTypeException
        End Select
    Else
        Set Parser = StringParser(IsSpecialArgument:=IsSpecialArgument)
    End If
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
Public Function File(Optional Path As String = "") As File
    Dim f As New File
    Call f.Initialize(Path)
    Set File = f
End Function
Public Function Registry(Optional Key As String = "") As Registry
    Dim R As New Registry
    Call R.Initialize(Key)
    Set Registry = R
End Function
Public Function Timer(Optional CallBack As Method, Optional Interval As Long = 1000, Optional InitialDelay As Long = 0, Optional StartNow As Boolean = False, Optional hWndParent As Long = 0) As Timer
    Dim t As New Timer
    Call t.Initialize(CallBack, Interval, InitialDelay, StartNow, hWndParent)
    Set Timer = t
End Function
Public Function MintAPITimer(Optional CallBack As Method, Optional Interval As Long = 1000, Optional InitialDelay As Long = 0, Optional StartNow As Boolean = False, Optional hWndParent As Long = 0) As Timer
    Dim t As New Timer
    Call t.Initialize(CallBack, Interval, InitialDelay, StartNow, hWndParent)
    Set MintAPITimer = t
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
    Call t.Initialize(0, 0)
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
Public Function Configuration(Optional Arguments) As Configuration
    Dim Conf As New Configuration
    Call Conf.Initialize(Arguments)
    Set Configuration = Conf
End Function
Public Function CurrentProcess() As Process
    Dim p As New Process
    Call p.Initialize(0, 0)
    Set CurrentProcess = p
End Function
Public Function Language(Optional Path As String = "") As Language
    Dim l As New Language
    Call l.Initialize(Path)
    Set Language = l
End Function
Public Function NoteBuffer(Optional Path As String = "") As NoteBuffer
    Dim nb As New NoteBuffer
    Call nb.Initialize(Path)
    Set NoteBuffer = nb
End Function
Public Function MIDIOutput() As MIDIOutput
    Dim MIDIO As New MIDIOutput
    Call MIDIO.Initialize
    Set MIDIOutput = MIDIO
End Function
Public Function MIDIInput() As MIDIInput
    Dim MIDII As New MIDIInput
    Call MIDII.Initialize
    Set MIDIOutput = MIDII
End Function

Public Function ArgumentList(Arguments() As Argument) As ArgumentList
    Dim al As New ArgumentList
    Call al.Initialize(Arguments)
    Set ArgumentList = al
End Function
Public Function List(ParamArray Items() As Variant) As List
    Dim l As New List, clonedItems() As Variant
    clonedItems = Items
    Call l.Initialize(clonedItems)
    Set List = l
End Function
Public Function Vector(Size As Long, vType As VbVarType) As Vector
    Dim vec As New Vector
    Call vec.Initialize(Size, vType)
    Set Vector = vec
End Function

Public Function Algorithms(Optional Arguments) As Algorithms
    Dim Alg As New Algorithms
    Call Alg.Initialize(Arguments)
    Set Algorithms = Alg
End Function
Public Function Alg(Optional Arguments) As Algorithms
    Dim al As New Algorithms
    Call al.Initialize(Arguments)
    Set Alg = al
End Function
Public Function ExAPI(Optional Arguments) As ExAPI
    Dim exa As New ExAPI
    Call exa.Initialize(Arguments)
    Set ExAPI = exa
End Function
Public Function Console() As Console
    Dim obj As Object
    Set obj = mod_application.appOutput(False)
    If Not obj Is Nothing Then
        If TypeOf obj Is Console Then
            Set Console = obj
            Exit Function
        End If
    End If
    Set Console = New Console
    Call Console.Initialize
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
Public Function Pixmap(Optional Source) As Pixmap
    Dim p As New Pixmap
    Call p.Initialize(Source)
    Set Pixmap = p
End Function
Public Function Graphics(Optional obj) As Graphics

End Function
Public Function CreateGraphicsOf(obj) As Graphics

End Function
Public Function GraphicsOf(obj) As Graphics: Set GraphicsOf = CreateGraphicsOf(obj): End Function
'==========================================================================================
'------------------------------------------------------------------------------------------
'..........................................................................................
'
'==========================================================================================
'== Object Buffer API Section =============================================================
'==========================================================================================
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
Public Function CreateEventNamesList(strEvents() As String) As Object
    Dim arr() As Variant
    Dim strEventsCount As Long
    strEventsCount = ArraySize(strEvents)
    If strEventsCount > 0 Then
        ReDim arr(strEventsCount - 1)
        Dim i As Long, xCounter As Long
        For i = LBound(strEvents) To UBound(strEvents)
            arr(xCounter) = CStr(strEvents(i))
            xCounter = xCounter + 1
        Next
    End If
    Set CreateEventNamesList = CreateObjectBufferC(OBJECTBUFFER_EVENTNAMES_LIST, arr)
End Function
Public Function CreateHandleObject(Handle As Long) As Object
    Set CreateHandleObject = CreateObjectBuffer(OBJECTBUFFER_HANDLE, Handle)
End Function
Public Function CreateHDCObject(hdc As Long) As Object
    Set CreateHDCObject = CreateObjectBuffer(OBJECTBUFFER_DISPLAY_CONTEXT, hdc)
End Function
Public Function CreateRenderTargetObject(target As Variant) As Object
    Set CreateRenderTargetObject = CreateObjectBuffer(OBJECTBUFFER_RENDERTARGET, target)
End Function
'!UNSAFE METHOD.
Public Function MintAPIType(TypeRefrence As Long, TypeLength As Long) As Object
    Set MintAPIType = CreateObjectBuffer(OBJECTBUFFER_MINTLOCALTYPE, TypeRefrence, TypeLength)
End Function
Public Function Inherit(Optional ByVal InheritFrom As Object = Nothing) As Object
    Set Inherit = CreateObjectBuffer(OBJECTBUFFER_INHERIT, InheritFrom)
End Function
Public Function DesktopWindow() As Object
    Set DesktopWindow = CreateHandleObject(uiMethods.GetDesktopWindowHandle)
End Function


'==========================================================================================
'------------------------------------------------------------------------------------------
'..........................................................................................

'==========================================================================================
'== Graphics API Section ==================================================================
'==========================================================================================

Public Function Color(ARGBValue As Long) As Color
    Color.Color = ARGBValue And &HFFFFFF
    Color.Alpha = ARGBValue And &HFF000000
End Function

'==========================================================================================
'------------------------------------------------------------------------------------------
'..........................................................................................

'==========================================================================================
'== MintAPI Methods Section ===============================================================
'==========================================================================================

Public Sub About()
    Call modMain.AboutMintAPI
End Sub

Public Function GetStreamOf(Stream) As ITargetStream
    If Not (TypeOf Stream Is ITargetStream) Then throw InvalidArgumentTypeException, Me
    Set GetStreamOf = Stream
End Function
Public Function GetDataOf(Data) As IData
    If Not (TypeOf Data Is IData) Then throw InvalidArgumentTypeException, Me
    Set GetDataOf = Data
End Function
Public Function GetCloneOf(ClonableClass As Object) As Object
    If ClonableClass Is Nothing Then throw ArgumentNullException, Me
    If Not TypeOf ClonableClass Is IClassClonable Then throw InvalidArgumentTypeException, Me
    Dim clnablObj As IClassClonable
    Set clnablObj = ClonableClass
    Set GetCloneOf = clnablObj.Clone
End Function
Public Function GetTexOf(TexableClass As Object) As Object
    If TexableClass Is Nothing Then throw ArgumentNullException, Me
    If Not TypeOf TexableClass Is IClassTexer Then throw InvalidArgumentTypeException, Me
    Dim texerObj As IClassTexer
    Set texerObj = TexableClass
    Set GetTexOf = CreateObjectBuffer(OBJECTBUFFER_TEXEDOBJECT, texerObj.toByteArray)
End Function

Public Function LastDllError() As Long
    LastDllError = API_GetLastError
End Function

Public Function MintAPIVersionString() As String
    MintAPIVersionString = APP_VERSIONSTRING
End Function
Public Function MintAPIMajorVersion() As Long
    MintAPIMajorVersion = App.Major
End Function
Public Function MintAPIMinorVersion() As Long
    MintAPIMinorVersion = App.Minor
End Function
Public Function MintAPIRevision() As Long
    MintAPIRevision = App.Revision
End Function
Public Function CheckMintAPIVersion(MajorVersion As Long, Optional MinorVersion As Long = -1, Optional Revision As Long = -1, Optional CheckVersionTag As String = "", Optional CheckOneOfVersionTags As String = "") As Boolean
    If MajorVersion > App.Major Then Exit Function
    If MinorVersion <> -1 Then _
        If MinorVersion > App.Minor Then Exit Function
    If Revision <> -1 Then _
        If Revision > App.Revision Then Exit Function

    If (Trim(CheckVersionTag) = "") And (Trim(CheckOneOfVersionTags) = "") Then
        CheckMintAPIVersion = True
    Else
        CheckMintAPIVersion = CheckMintAPIVersionTag(CheckVersionTag, CheckOneOfVersionTags)
    End If
End Function
Public Function MintAPIDllPath() As String
    Dim strName As String
    strName = App.EXEName
    'If Right(Trim(strName), 4) <> ".dll" Then strName = strName & ".dll"
    MintAPIDllPath = Directory.ConcatPath(App.Path, strName)
End Function
Public Function MintAPIPath() As String
    MintAPIPath = App.Path
End Function
Public Function CheckMintAPIVersionTag(strVersionTags As String, Optional strConditionalVersionTags As String = "") As Boolean
    Dim strVersionTagsTRIMMED As String
    Dim strConditionalVersionTagsTRIMMED As String
    Dim APP_VERSIONTAGSTRIMMED As String
    strVersionTagsTRIMMED = Trim(strVersionTags)
    strConditionalVersionTagsTRIMMED = Trim(strConditionalVersionTags)
    If (strVersionTagsTRIMMED = "") And (strConditionalVersionTagsTRIMMED = "") Then
        APP_VERSIONTAGSTRIMMED = Trim(APP_VERSIONTAGS)
        If APP_VERSIONTAGS = "" Then
            CheckMintAPIVersionTag = True
        Else
            CheckMintAPIVersionTag = False
        End If
        Exit Function
    End If
    Dim str1() As String
    Dim str2() As String
    Dim sVersionTags() As String
    APP_VERSIONTAGSTRIMMED = Trim(APP_VERSIONTAGS)
    str1 = Split(Trim(strVersionTags), " ")
    str2 = Split(Trim(strConditionalVersionTags), " ")
    sVersionTags = Split(APP_VERSIONTAGSTRIMMED, " ")
    Dim str1Count As Long, str2Count As Long, sVTCount As Long
    str1Count = ArraySize(str1)
    str2Count = ArraySize(str2)
    sVTCount = ArraySize(sVersionTags)

    If str1Count <> sVTCount Then _
        If str1Count > 0 Then Exit Function

    Dim i As Long, j As Long, str2Found As Long
    For i = 0 To str1Count - 1
        For j = 0 To sVTCount - 1
            If str1(i) <> sVersionTags(j) Then
                CheckMintAPIVersionTag = False
                Exit Function
            End If
        Next
    Next
    For i = 0 To str2Count - 1
        For j = 0 To sVTCount - 1
            If str2(i) = sVersionTags(j) Then
                CheckMintAPIVersionTag = True
                str2Found = str2Found + 1
                Exit Function
            End If
        Next
    Next
    If str1Count <= 0 Then
        If str2Found > 0 Then
            CheckMintAPIVersionTag = True
        Else
            CheckMintAPIVersionTag = False
        End If
    Else
        CheckMintAPIVersionTag = True
    End If
End Function
'==========================================================================================
'------------------------------------------------------------------------------------------
'..........................................................................................

'==========================================================================================
'== Signal & Slot Connection Section ==================================================================
'==========================================================================================
Public Function Signal(Name As String) As Object
    Set Signal = CreateObjectBuffer(OBJECTBUFFER_SIGNAL, Name)
End Function
Public Function Slot(Name As String, SlotObject As Object) As Object
    Set Slot = CreateObjectBuffer(OBJECTBUFFER_SLOT, Name, SlotObject)
End Function
Public Function mSlot(Method As Method) As Object
    Set mSlot = CreateObjectBuffer(OBJECTBUFFER_SLOT, Method)
End Function
Public Function dSlot(Address As Long) As Object
    Set dSlot = CreateObjectBuffer(OBJECTBUFFER_SLOT, Address)
End Function
Public Function cSlot(CallBack As ICallBack) As Object
    Set cSlot = CreateObjectBuffer(OBJECTBUFFER_SLOT, CallBack)
End Function
Public Sub Connect( _
        Signal As Object, SignalObject As ISupportSignal, _
        Slot As Object)
    'If Not TypeOf SlotObject Is ISupportSlot Then _
        throw InvalidArgumentTypeException("at API.Connect() , SlotObject type must be as ISupportSlot.")

    If SignalObject Is Nothing Then throw ArgumentNullException

    Dim SignalOB As ObjectBuffer, SlotOB As ObjectBuffer
    Set SignalOB = Signal
    Call SignalOB.thrownot(OBJECTBUFFER_SIGNAL, 1)
    Set SlotOB = Slot
    If Not SlotOB.IsSlot Then throw InvalidArgumentTypeException("Invalid slot.")
    Dim Emitter As SignalEmitter
    Set Emitter = SignalObject.Emitter
    Call Emitter.Append(SignalOB.First, SlotOB)
End Sub
Public Sub Disconnect( _
        Signal As Object, SignalObject As ISupportSignal, _
        Slot As Object)
    'If Not TypeOf SlotObject Is ISupportSlot Then _
        throw InvalidArgumentTypeException("at API.Connect() , SlotObject type must be as ISupportSlot.")

    If SignalObject Is Nothing Then throw ArgumentNullException

    Dim SignalOB As ObjectBuffer, SlotOB As ObjectBuffer
    Set SignalOB = Signal
    Call SignalOB.thrownot(OBJECTBUFFER_SIGNAL, 1)
    Set SlotOB = Slot
    If Not SlotOB.IsSlot Then throw InvalidArgumentTypeException("Invalid slot.")
    Dim Emitter As SignalEmitter
    Set Emitter = SignalObject.Emitter
    Call Emitter.Remove(SignalOB.First, SlotOB)
End Sub
Public Sub Emit(Signal, SignalObject As ISupportSignal)
    Call SignalObject.Emitter.Emit(Signal)
End Sub
'==========================================================================================
'------------------------------------------------------------------------------------------
'..........................................................................................

'==========================================================================================
'== Bonus Section ===============================================================
'==========================================================================================

' Description : IncVar(x) is the same as the C operator ++x
Public Function IncVar(ByRef Value As Variant, Optional ByVal Amount As Variant = 1) As Variant
    Value = Value + Amount
    IncVar = Value
End Function
' Description : DecVar(x) is the same as the C operator --x
Public Function DecVar(ByRef Value As Variant, Optional ByVal Amount As Variant = 1) As Variant
    Value = Value - Amount
    DecVar = Value
End Function
' Description : VarInc(x) is the same as the C operator x++
Public Function VarInc(ByRef Value As Variant, Optional ByVal Amount As Variant = 1) As Variant
    VarInc = Value
    Value = Value + Amount
End Function
' Description : VarDec(x) is the same as the C operator x--
Public Function VarDec(ByRef Value As Variant, Optional ByVal Amount As Variant = 1) As Variant
    VarDec = Value
    Value = Value - Amount
End Function
' Description : assigns a value and returns it
Public Function LetVar(ByRef Variable As Variant, Value As Variant) As Variant
    If IsObject(Value) Then
        Set Variable = Value
        Set LetVar = Variable
    Else
            Variable = Value
            LetVar = Variable
    End If
End Function
' Description : returns the variable value and, after that, assigns new value
Public Function VarLet(ByRef Variable As Variant, Value As Variant) As Variant
    If IsObject(Value) Then
        Set VarLet = Variable
        Set Variable = Value
    Else
            VarLet = Variable
            Variable = Value
    End If

End Function
'==========================================================================================
'------------------------------------------------------------------------------------------
'..........................................................................................


Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property
