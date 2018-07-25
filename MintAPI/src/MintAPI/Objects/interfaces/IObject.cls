VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"MetaObject"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

Option Base 0
Option Explicit
Const CLASSID As String = "IDevice"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING

''<class interface>
''<summary>Provide a basic class for entire framework.</summary>
''<version 0.1/>
''<ideas>Me , .net Framework , Qt</ideas>
''</class>

''<summary>
''  Used to automatic class provider to construct the object.
''</summary>
''<params>
''<param name="CInfo" type="ObjectConstructionInfo" byval in>
''  An object contains required info to construct the object.
''</param>
''<samples>
''  Dim myFile As Object
''  Set myFile = ClassProvider.ProvideClass("com.mintapi.file", ArgumentList("C:\t.txt"))
''   ' The second argument may pass to IObject::Initialize() method().
''  Call myFile.Open
''  ....
''</samples>
''<see>IProvider,Provider</see>
Public Sub Construct(ByVal CInfo As ObjectConstructionInfo): End Sub

''<summary>
''  Returns an alias string to uniquely identify the object type.
''  It also can use slash to indicate the version of the class.
''  This property will use for class provider system.
''</summary>
''<remarks>
''  It can be used combination of alias and version together in class provider.
''</remarks>
''<sample>com.java... , com.mintapi.language/0.1</sample>
''<obsolete>
'Public Property Get Alias() As String: End Property

''<summary>
''  A string representing the object default value or if
''  Target object does'nt have default value returns a string
''  Representing the object and its namespace.
''</summary>
''<retval type="string">A string to represent the object.</retval>
Public Function ToString() As String: End Function

''<summary>
''  Returns a pseudo-unique hashed 4 bytes number to identify this object instance.
''</summary>
''<retval type="int32">The object unique hash code.</retval>
Public Function GetHashCode() As Long: End Function

''<summary>
''  Determines if passed argument is equal to the object instance.
''</summary>
''<params>
''<param name="Value" type="variant">
''</params>
''<retval type="bool">A boolean indicating the equality of this instance versus the passed in value.</retval>
Public Function Equals(Value As Variant) As Boolean: End Function

''<summary>
''  Return a metaobject for the object to control the object. null if object does'nt have meta object.
''</summary>
''<retval type="MintAPI.MetaObject">A meta object for the object instance.</retval>
Public Function MetaObject() As MetaObject: End Function

''<summary>
''  Provides a mechanism to query information from IObject derived classes.
''</summary>
''<remarks>
''  It must add the responses to return the ObjectQR passed as input argument and return it.
''</remarks>
''<sample>
''  stdout Objects.QueryInformation(myClass, "private_name", args).Responses(0)
''</sample>
Public Function Stub(ByVal Proxy As ObjectProxyInfo) As Boolean: End Function
