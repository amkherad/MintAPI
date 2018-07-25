VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Version"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"MetaObject"
'@PROJECT_LICENSE

''This class is taken from VBCorLib.dll by Kelly Ethridge

''<class name="Version" serializable tags="immutable">
''  <summary>
''    Represents a version number.
''  </summary>
''  <constructors>
''      <constructor name="" friend>
''          Some constructor description.
''      </constructor>
''  </constructors>
''  <using></using>
''  <idea>Me , VBCorLib</idea>
''  <sample>
''      Dim v As Version
''      Set v = Version.Parse("1.0.0.1001")
''<br>
''      stdout v.ToString
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
Const CLASSID As String = "Version" '#DO NOT REMOVE IT
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID '#DO NOT REMOVE IT
Const CLASSVERSIONSTRING As String = "0.1" '#DO NOT REMOVE IT
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING '#DO NOT REMOVE IT

Const SIGNAL_TERMINATING                As String = "terminating"
Const SLOT_METAOBJECT                   As String = "metaobject"

Const PROPERTY_MAJOR                    As String = "Major"
Const PROPERTY_MINOR                    As String = "Minor"
Const PROPERTY_BUILD                    As String = "Build"
Const PROPERTY_REVISION                 As String = "Revision"

Const VERSION_UNUSED                    As Long = -1
Const DEFAULT_MAJOR                     As Long = 0
Const DEFAULT_MINOR                     As Long = 0
Const DEFAULT_BUILD                     As Long = VERSION_UNUSED
Const DEFAULT_REVISION                  As Long = VERSION_UNUSED
'by Kelly Ethridge
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section API/Type Declarations>
'

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

Private mMajor      As Long
Private mMinor      As Long
Private mBuild      As Long
Private mRevision   As Long
'by Kelly Ethridge
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
'    Dim ArgsCount As Long
'    ArgsCount = Arguments.Count
'    If ArgsCount = 1 Then
'        Call Constructor1(Arguments(0))
'    ElseIf ArgsCount > 1 Then
'On Error GoTo CatchedError
'        Dim vMajor As Long, vMinor As Long, vBuild As Long, vRevision As Long
'        vMajor = CLng(Arguments.GetValue(0, 0))
'        vMinor = CLng(Arguments.GetValue(1, 0))
'        vBuild = CLng(Arguments.GetValue(2, 0))
'        vRevision = CLng(Arguments.GetValue(3, 0))
'CatchedError:
'    End If

    This_Constructed = True
End Sub
Friend Sub Constructor0(ByVal Major As Long, ByVal Minor As Long, ByRef vBuild As Variant, ByRef vRevision As Variant) 'by Kelly Ethridge
    If Major < 0 Or Minor < 0 Then _
        throw Exps.IndexOutOfRangeException(Mtr("Version's parameters must be equal to or greater than zero."))
    
    If Not IsMissing(vBuild) Then
        mBuild = vBuild
        If mBuild < 0 Then throw Exps.IndexOutOfRangeException(Mtr("Version's parameters must be equal to or greater than zero."))
    End If
    If Not IsMissing(vRevision) Then
        If mBuild < 0 Then throw Exps.IndexOutOfRangeException(Mtr("Cannot include a revision without a build value."))
        mRevision = vRevision
        If mRevision < 0 Then throw Exps.IndexOutOfRangeException(Mtr("Version's parameters must be equal to or greater than zero."))
    End If
    
    mMajor = Major
    mMinor = Minor
    
    This_Constructed = True
End Sub
Friend Sub Constructor1(ByRef vStr As String)
    Dim Parts() As String
    Dim ub As Long
    
    Parts = Split(vStr, ".")
    ub = UBound(Parts)
    If ub < 1 Or ub > 3 Then _
        throw Exps.InvalidArgumentException("String must contain 2 to 4 fields.")
    mMajor = CLng(Parts(0))
    mMinor = CLng(Parts(1))
    If mMajor < 0 Or mMinor < 0 Then _
        throw Exps.IndexOutOfRangeException("Version's parameters must be equal to or greater than zero.")
    
    If ub > 1 Then
        mBuild = CLng(Parts(2))
        If mBuild < 0 Then throw Exps.IndexOutOfRangeException("Version's parameters must be equal to or greater than zero.")
        If ub > 2 Then
            mRevision = CLng(Parts(3))
            If mRevision < 0 Then throw Exps.IndexOutOfRangeException("Version's parameters must be equal to or greater than zero.")
        End If
    End If
    
    This_Constructed = True
End Sub
Friend Sub Constructor2(ByVal lngVersion As Long)
    
End Sub
Friend Sub ConstructorCopy(ByVal Version As Version)
    mMajor = Version.Major
    mMinor = Version.Minor
    mBuild = Version.Build
    mRevision = Version.Revision
    
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

''<summary>Returns the major version number of this instance.</summary>
''<retval>The major version number. This is a non-negative number.</retval>
Public Property Get Major() As Long
    Major = mMajor
End Property

''<summary>Returns the minor version number of this instance.</summary>
''<retval>The minor version number. This is a non-negative number.</retval>
Public Property Get Minor() As Long
    Minor = mMinor
End Property

''<summary>Returns the build version number of this instance.</summary>
''<retval>The build version number.</retval>
''<remarks>
''  The build version is an optional tracking number. If the number
''  was not used when this instance was created, then -1 is returned.
''</remarks>
Public Property Get Build() As Long
    Build = mBuild
End Property

''<summary>Returns the revision version number of this instance.</summary>
''<retval>The revision version number.</retval>
''<remarks>
''  The build version is an optional tracking number. If the number
''  was not used when this instance was created, then -1 is returned.
''</remarks>
Public Property Get Revision() As Long
    Revision = mRevision
End Property

''<summary>Returns a minor revision value.</summary>
''<retval>Minor revision number.</retval>
''<remarks>
''  The minor revision is the lower 16-bits of the
''  standard revision value.
''</remarks>
Public Property Get MinorRevision() As Long
    MinorRevision = mRevision And &HFFFF
End Property

''<summary>Returns the major revision value.</summary>
''<retval>Major revision number.</retval>
''<remarks>
''  The major revision number is the upper 16-bits
''  of the standard revision value.
''</remarks>
Public Property Get MajorRevision() As Long
    MajorRevision = ShiftRight(mRevision, 16)
End Property

''<summary>Returns a cloned object of this instance.</summary>
''<retval>The cloned version of this instanced.</retval>
Public Function Clone() As Version
    Set Clone = New Version
    Call Clone.ConstructorCopy(Me)
End Function

''<summary>Compares this instance to a passed in value.</summary>
''<params><param Name="Value">The value to compare this instance against.</param></params>
''<retval>Value indicating the relation of this instance to the passed in value.</retval>
''<remarks>
''  A negative number indicates this instance is less than the value.
''  A positive number indicates this instance is greater than the value.
''  Zero indicates this instance is equal to the value.
''</remarks>
Public Function CompareTo(ByRef Value As Variant) As CompareResults
    Select Case VarType(Value)
        Case VBObject
            If Value Is Nothing Then: CompareTo = 1: Exit Function
            If Not TypeOf Value Is Version Then throw Exps.InvalidArgumentException("Value must of type Version.")

            Dim V As Version
            Set V = Value
            If Not EqualFields(mMajor, V.Major, CompareTo) Then Exit Function
            If Not EqualFields(mMinor, V.Minor, CompareTo) Then Exit Function
            If Not EqualFields(mBuild, V.Build, CompareTo) Then Exit Function
            If Not EqualFields(mRevision, V.Revision, CompareTo) Then Exit Function

        Case vbNull, vbEmpty, vbError
            CompareTo = 1
        Case vbString
            CompareTo = MintAPI.Version.Parse(CStr(Value)).CompareTo(Me)
        Case Else
            throw Exps.InvalidArgumentException("Value must of type Version.")
    End Select
End Function

Public Function EqualTo(ByRef Value As Version) As Boolean
    If Value Is Nothing Then Exit Function
    EqualTo = (CompareTo(Value) = crEqual)
End Function

Public Function LessThan(ByRef Value As Version) As Boolean
    If Value Is Nothing Then _
        throw Exps.ArgumentNullException(Mtr("Cannot compare a Version to Nothing."))
    
    LessThan = (CompareTo(Value) < crEqual)
End Function

Public Function LessThanOrEqualTo(ByRef Value As Version) As Boolean
    If Value Is Nothing Then _
        throw Exps.ArgumentNullException(Mtr("Cannot compare a Version to Nothing."))
    
    LessThanOrEqualTo = (CompareTo(Value) <= crEqual)
End Function

Public Function GreaterThan(ByRef Value As Version) As Boolean
    GreaterThan = Not LessThanOrEqualTo(Value)
End Function

Public Function GreaterThanOrEqualTo(ByRef Value As Version) As Boolean
    GreaterThanOrEqualTo = Not LessThan(Value)
End Function

Public Function NotEqualTo(ByRef Value As Version) As Boolean
    NotEqualTo = Not EqualTo(Value)
End Function

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
Public Function Equals(ByRef Value As Variant) As Boolean
    If IsObject(Value) Then
        If Value Is Nothing Then Exit Function
        
        Dim V As Version
        Set V = Version
        If TypeOf V Is Version Then
            Equals = (CompareTo(V) = crEqual)
        End If
    End If
End Function
Public Function ToString(Optional ByRef vFieldCount As Variant) As String
    Dim RetVal As String
    Dim FieldCount As Long
    Dim MaxField As Long
    
    If IsMissing(vFieldCount) Then
        FieldCount = -1
    Else
        FieldCount = vFieldCount
        MaxField = GetMaxFieldCount
        'If FieldCount < 0 Or FieldCount > MaxField Then throw Exps.Cor.NewArgumentOutOfRangeException(Environment.GetResourceString(ArgumentOutOfRange_VersionFieldCount, MaxField), "vFieldCount", FieldCount)
    End If
    
    If FieldCount <> 0 Then RetVal = mMajor
    If FieldCount < 0 Or FieldCount > 1 Then RetVal = RetVal & "." & mMinor
    If mBuild >= 0 And (FieldCount < 0 Or FieldCount > 2) Then
        RetVal = RetVal & "." & mBuild
        If mRevision >= 0 And (FieldCount < 0 Or FieldCount > 3) Then RetVal = RetVal & "." & mRevision
    End If
    ToString = RetVal
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
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Private helpers>
'
Private Sub Class_InitProperties()
    mBuild = VERSION_UNUSED
    mRevision = VERSION_UNUSED
End Sub
Private Sub Class_Terminate()
    If Not myMetaObject Is Nothing Then
        Signal.Emit Me, SIGNAL_TERMINATING
        Call myMetaObject.Dispose
        Set myMetaObject = Nothing
    End If
End Sub
Private Sub Class_ReadProperties(PropBag As PropertyBag)
    With PropBag
        mMajor = .ReadProperty(PROPERTY_MAJOR, DEFAULT_MAJOR)
        mMinor = .ReadProperty(PROPERTY_MINOR, DEFAULT_MINOR)
        mBuild = .ReadProperty(PROPERTY_BUILD, DEFAULT_BUILD)
        mRevision = .ReadProperty(PROPERTY_REVISION, DEFAULT_REVISION)
    End With
End Sub
Private Sub Class_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty(PROPERTY_MAJOR, mMajor)
        Call .WriteProperty(PROPERTY_MINOR, mMinor)
        Call .WriteProperty(PROPERTY_BUILD, mBuild)
        Call .WriteProperty(PROPERTY_REVISION, mRevision)
    End With
End Sub

Private Function GetMaxFieldCount() As Long 'Kelly Ethridge
    If mBuild < 0 Then
        GetMaxFieldCount = 2
    ElseIf mRevision < 0 Then
        GetMaxFieldCount = 3
    Else
        GetMaxFieldCount = 4
    End If
End Function
Private Function EqualFields(ByVal MyField As Long, ByVal TheirField As Long, ByRef RetVal As Long) As Boolean 'Kelly Ethridge
    If MyField < TheirField Then
        RetVal = -1
    ElseIf MyField > TheirField Then
        RetVal = 1
    Else
        RetVal = 0
        EqualFields = True
    End If
End Function
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
