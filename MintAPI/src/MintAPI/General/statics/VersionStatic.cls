VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VersionStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="VersionStatic" static tags="">
''  <summary>
''    Provides version static methods.
''  </summary>
''  <using>Version</using>
''  <idea>Me</idea>
''  <sample>
''      Dim v As Version
''      Set v = Version.Parse("1.0.0.1001")<br>
''<br>
''      Version.Parse("1.0.0.1001")
''  </sample>
''</class>

Option Base 0
Option Explicit

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "VersionStatic" '#DO NOT REMOVE IT
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID '#DO NOT REMOVE IT
Const CLASSVERSIONSTRING As String = "0.1" '#DO NOT REMOVE IT
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING '#DO NOT REMOVE IT

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
'<section Creators>
'
Public Function Create(ByVal Major As Long, ByVal Minor As Long, Optional ByVal Build As Long = 0, Optional ByVal Revision As Long = 0) As Version
    Set Create = New Version
    Call Create.Constructor0(Major, Minor, Build, Revision)
End Function
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

''<summary>Parses a string in the format of 0.0[.0][.0]</summary>
''<params><param name="vStr">The string to be parsed.</param></params>
''<retval>A version defined from the string.</retval>
''<remarks>Atleast 2 of the version numbers must be supplied, and
''all version numbers must be non-negative.</remarks>
Public Function Parse(ByVal vStr As String) As Version
    Set Parse = New Version
    Call Parse.Constructor1(vStr)
End Function

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------
