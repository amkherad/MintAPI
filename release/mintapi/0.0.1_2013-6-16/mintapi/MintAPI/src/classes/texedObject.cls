VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "texedObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Base 0
Option Explicit
Const CLASSID As String = "texedObject"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


Implements IAliasable


'Stringed format.===================================================
'===================================================================
'MClassTexerS0.1<classname,classalias/classversionstring>[[classparameters]]{ //ClassAlias used to load class using HostProvider.
' propertyname1:propertyvalue1;
' propertyname2:propertyvalue2;
'             ...
'             ...
'             ...
' propertynamen:propertyvaluen;
'}
'====================================================================
'ByteArray format.===================================================
'====================================================================
'MClassTexerB0.1<classname_DYN,classalias_DYN/classversionstring_8>classparameters_4
'propertyname1_DYN:propertyvalue1_DYN
'propertyname2_DYN:propertyvalue2_DYN
'            ...
'            ...
'            ...
'propertynamen_DYN:propertyvaluen_DYN
'====================================================================

Dim txObject As IClassTexer

Dim texString As String
Dim texBA As ByteArray
Dim IsStringed As Boolean
Dim IsByteArrayed As Boolean

Dim tx_header As String

Dim tx_EngineVersion As Long
Dim tx_Name As String
Dim tx_ORGTypeName As String
Dim tx_TypeName As String
Dim tx_TypeAlias As String
Dim tx_Version As Long
Dim tx_Arguments As Long

Friend Sub Initialize(texableObject As IClassTexer)
    If texableObject Is Nothing Then throw InvalidStatusException("Set a texable object first.")
    Set txObject = texableObject
    tx_header = txObject.texHeader(tt_String)
    tx_ORGTypeName = VBA.TypeName(texableObject)
End Sub

Private Sub EnsureString()
    If Not IsStringed Then _
        texString = txObject.toString
End Sub
Private Sub EnsureBA()
    If Not IsByteArrayed Then _
        Set texBA = ByteArray(txObject.toByteArray)
End Sub
Private Sub EnsureObject()
    If txObject Is Nothing Then throw InvalidStatusException("Set a texable object first.")
End Sub

'Version of class texer engine.
Public Property Get TexerVersion() As Long
    If tx_EngineVersion <> 0 Then
        TexerVersion = tx_EngineVersion
    Else
        TexerVersion = CLASSTEXER_VERSION
    End If
End Property

Public Property Get Name() As String
EnsureObject
    Name = tx_Name
End Property
Public Property Get TypeName() As String
EnsureObject
    TypeName = tx_TypeName
End Property
Public Property Get TypeAlias() As String
EnsureObject
    TypeAlias = tx_TypeAlias
End Property
Public Property Get Version() As Long
EnsureObject
    Version = tx_Version
End Property
Public Property Get Arguments() As Long
EnsureObject
    Arguments = tx_Arguments
End Property

Public Function GetArguments()
    
End Function


Public Function toTexString() As String
EnsureObject
EnsureString
    toTexString = texString
End Function
Public Sub fromTexString(strTex As String)
EnsureObject
    Call txObject.fromString(strTex)
End Sub
Public Function toTexByteArray() As ByteArray
EnsureObject
EnsureBA
    Set toTexByteArray = texBA.Clone
End Function
Public Sub fromTexByteArray(baTex As ByteArray)
EnsureObject
    Call txObject.fromByteArray(baTex.constData)
End Sub

Public Function toXMLString() As String
EnsureObject
    Dim xml_Header As String, xml_Content As String, xml_Terminator As String
    
    xml_Header = "<object " & _
                 "name=""" & Name & """ " & _
                 "type=""" & TypeName & """ " & _
                 "orgtype=""" & tx_ORGTypeName & """ " & _
                 "alias=""" & TypeAlias & """ " & _
                 "version=""" & Version & """ " & _
                 "args=""" & Arguments & """ " & _
                 ">" & vbCrLf & _
                 "  <param type=""class""/>" & vbCrLf
                 

    xml_Content = "  content  "
    
    xml_Terminator = vbCrLf & "</object>"
    
    toXMLString = xml_Header & xml_Content & xml_Terminator
End Function
Public Function toSOAPString() As String
EnsureObject
    toSOAPString = "<!DOCTYPE>" & vbCrLf & toXMLString
End Function

Public Sub SaveToStream(targetStream As ITargetStream)
EnsureObject
    Call targetStream.outStream(toTexString)
End Sub

Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property
