VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OperatingSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Version"
Attribute VB_Ext_KEY = "Member1" ,"MetaObject"
'@PROJECT_LICENSE

'This class originally written by Kelly Ethridge at VBCorLib.

''<class name="OperatingSystem" tags="">
''  <summary>
''    Provides information about the currently running operating system.
''  </summary>
''  <constructors>
''      <constructor name="" friend>
''          Some constructor description.
''      </constructor>
''  </constructors>
''  <using>Version ICloneable</using>
''  <idea>Me , .net Framework , VBCorLib</idea>
''  <sample>
''      stdout OperatingSystem.ToString, endl
''  </sample>
''</class>

Option Base 0
Option Explicit

Implements IObject
Implements ICloneable

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "OperatingSystem" '#DO NOT REMOVE IT
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID '#DO NOT REMOVE IT
Const CLASSVERSIONSTRING As String = "0.1" '#DO NOT REMOVE IT
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING '#DO NOT REMOVE IT

Const SIGNAL_TERMINATING As String = "terminating"
Const SLOT_METAOBJECT As String = "metaobject"

Const PROPERTY_PLATFORMID As String = "PlatformID"
Const PROPERTY_VERSION As String = "Version"
Const PROPERTY_SERVICEPACK As String = "ServicePack"
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section API/Type Declarations>
'
Public Enum PlatformID 'MSDN
    Win32S = 0
    Win32Windows = 1
    WinNT = 2
    WinCE = 3
End Enum
Public Enum OSVersion
    OSDOS = 0
    OSWin3x = 1
    OSWin95 = 2
    OSWin98 = 3
    OSWinNT = 4
    OSWinNT3X = OSWinNT
    OSWinNT40 = OSWinNT
    OSWin2000 = 5
    OSWinXP = OSWin2000
    OSWinServer2003 = OSWinXP
    OSWinVista = 6
    OSWin7 = 7
    OSWinServer2008 = OSWin7
    OSWin8 = 8
    'OSWin9 = 9
    OSHigher = &H80
    OSUnknown = &H7F
End Enum

'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Variables/Events>
'
Dim This_Constructed As Boolean
Dim myMetaObject As MetaObject

Private p_PlatformID     As PlatformID
Private p_Version        As Version
Private p_Loaded         As Boolean
Private p_ServicePack    As String

'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Constructors>
'
''<summary>
''  Basic constructor to use by MintAPI class provider,
''  When creating new instance of this class.
''</summary>
''<params>
''  <param name="Arguments">
''      A list of construction arguments.
''  </param>
''</params>
''<cargs order="0" count="1">
''  <carg type="long" name="ID"></carg>
''</cargs>
Friend Sub Constructor(ByVal CInfo As ObjectConstructionInfo)
    
    This_Constructed = True
End Sub
Friend Sub Constructor0(ByVal Platform As PlatformID, ByVal Version As Version, ByVal ServicePack As String)
    p_PlatformID = Platform
    Set p_Version = Version
    p_ServicePack = ServicePack
    p_Loaded = True
    
    This_Constructed = True
End Sub
Friend Sub ConstructorCopy(ByVal OS As OperatingSystem)
    p_PlatformID = OS.Platform
    Set p_Version = OS.Version
    p_ServicePack = OS.ServicePack
    p_Loaded = True
    
    This_Constructed = True
End Sub
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Class Public Body Members>
'
'*********************************************
'*********************************************

''<summary>Returns the service pack for the current operating system.</summary>
''<remarks>An empty string is always returned on Windows 95, 98, and ME.</remarks>
Public Property Get ServicePack() As String: ENSURE_Loaded
    ServicePack = p_ServicePack
End Property

''<summary>Returns the operating system type currently running on this machine.</summary>
''<retval>The current operating system running.</retval>
Public Property Get Platform() As PlatformID: ENSURE_Loaded
    Platform = p_PlatformID
End Property

''<summary>Returns the version of the current operating system.</summary>
''<retval>The operating system version.</retval>
Public Property Get Version() As Version: ENSURE_Loaded
    Set Version = p_Version
End Property

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Basic Members>
'
Public Function Equals(Value As Variant) As Boolean
    Equals = Objects.Equals(Me, Value)
End Function
Public Function ToString() As String: ENSURE_Loaded
    Dim RetVal As String
    Select Case p_PlatformID
        Case WinNT
            If p_Version.Major = 5 Then
                Select Case p_Version.Minor
                    Case 0:     RetVal = "Microsoft Windows 2000"
                    Case 1:     RetVal = "Microsoft Windows XP"
                    Case 2:     RetVal = "Microsoft Windows Server 2003 family"
                    Case Else:  RetVal = "Microsoft Windows NT"
                End Select
            Else
                RetVal = "Microsoft Windows NT"
            End If
        Case Win32S
            RetVal = "Microsoft Win32s"
        Case Win32Windows
            Select Case p_Version.Minor
                Case 0:     RetVal = "Microsoft Windows 95"
                Case 10:    RetVal = "Micorosft Windows 98"
                Case 90:    RetVal = "Micorsoft Windows ME"
                Case Else:  RetVal = "Microsoft Windows32"
            End Select
        Case WinCE
            RetVal = "Microsoft WinCE"
    End Select
    
    If Len(p_ServicePack) > 0 Then
        ToString = RetVal & " " & p_Version.ToString(3) & " " & p_ServicePack
    Else
        ToString = RetVal & " " & p_Version.ToString
    End If
End Function
Public Function GetHashCode() As Long
    GetHashCode = Objects.GetHashCode(Me)
End Function
Public Function MetaObject() As MetaObject
    If myMetaObject Is Nothing Then
        Call Objects.MetaObject.Setup(myMetaObject, Me, True, CLASSALIASSTRING)
    End If
    Set MetaObject = myMetaObject
End Function

Public Function Clone() As OperatingSystem
    Dim OS As New OperatingSystem
    Call OS.ConstructorCopy(Me)
End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Private helpers>
'
Private Sub Class_Terminate()
    If Not myMetaObject Is Nothing Then
        Signal.Emit Me, SIGNAL_TERMINATING
        Call myMetaObject.Dispose
        Set myMetaObject = Nothing
    End If
End Sub
Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With PropBag
        p_PlatformID = .ReadProperty(PROPERTY_PLATFORMID, 0)
        Set p_Version = .ReadProperty(PROPERTY_VERSION, Nothing)
        p_ServicePack = .ReadProperty(PROPERTY_SERVICEPACK, vbNullString)
    End With
    p_Loaded = True
End Sub
Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty(PROPERTY_PLATFORMID, p_PlatformID)
        Call .WriteProperty(PROPERTY_VERSION, p_Version)
        Call .WriteProperty(PROPERTY_SERVICEPACK, p_ServicePack)
    End With
End Sub

Private Sub ENSURE_Loaded() 'from VBCorLib VerifyLoaded
    If p_Loaded Then Exit Sub
    
    Dim Info As API_OSVERSIONINFOA
    Info.dwOSVersionInfoSize = Len(Info)
    If API_GetVersionExUnicode(Info) = 0 Then _
        throw Exps.InvalidCallException("Could not load Operating System information.")

    ' We are on an NT platform so we can get extra information
    If Info.dwPlatformId = WinNT Then
        Dim InfoEX As API_OSVERSIONINFOEXA
        InfoEX.dwOSVersionInfoSize = Len(InfoEX)
        If API_GetVersionExUnicode(InfoEX) = 0 Then _
            throw Exps.InvalidOperationException("Could not load Operating System information.")

        Dim ServicePack As String
        ServicePack = StrConv(InfoEX.szCSDVersion, vbUnicode)
        ServicePack = API_SysAllocString(StrPtr(ServicePack))
    End If
    Call Constructor0(Info.dwPlatformId, MintAPI.Version.Create(Info.dwMajorVersion, Info.dwMinorVersion, Info.dwBuildNumber, ShiftLeft(InfoEX.wServicePackMajor, 16) Or InfoEX.wServicePackMinor), ServicePack)
End Sub
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section IObject Implementation>
'
Private Sub IObject_Construct(ByVal CInfo As ObjectConstructionInfo): Call Constructor(CInfo): End Sub
Private Function IObject_Equals(Value As Variant) As Boolean: IObject_Equals = Equals(Value): End Function
Private Function IObject_GetHashCode() As Long: IObject_GetHashCode = GetHashCode: End Function
Private Function IObject_MetaObject() As MetaObject: Set IObject_MetaObject = MetaObject: End Function
Private Function IObject_ToString() As String: IObject_ToString = ToString: End Function
Private Function IObject_Stub(ByVal ProxyInfo As ObjectProxyInfo) As Boolean
    
    IObject_Stub = False
End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section ICloneable Implementations>
'
Private Function ICloneable_Clone() As Object: Set ICloneable_Clone = Clone: End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------
