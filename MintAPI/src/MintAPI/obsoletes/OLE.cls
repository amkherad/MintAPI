VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OLE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Version"
'@PROJECT_LICENSE

''<class name="OLE" static tags="">
''  <summary>
''    Provides OLE wrapper methods.
''  </summary>
''  <using></using>
''  <idea>Me</idea>
''  <sample>
''      OLE.RegisterActiveX("C:\myAxDll.dll")
''      OLE.RegisterActiveX("C:\myAxExe.exe")
''      OLE.RegisterActiveX("C:\myAxOcx.ocx")
''  </sample>
''</class>

Option Base 0
Option Explicit

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "OLE" '#DO NOT REMOVE IT
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

Public Function Initialize()
    Initialize = API_OleInitialize(ByVal vbNullPtr)
End Function

Public Function RegisterActiveX(ByVal Path As String) As Long
    
End Function
Public Function UnRegisterActiveX(ByVal Path As String) As Long
    
End Function

Public Function CreateInstance(ByVal AliasString As String, Optional ByVal ServerName As String = "") As Object
    Set CreateInstance = VBA.CreateObject(AliasString, ServerName)
End Function
Public Function CreateInstanceFromFile(ByVal FilePath As String, ByVal ClassName As String) As Object
    Set CreateInstanceFromFile = VBA.GetObject(FilePath, ClassName)
End Function
Public Function CreateInstanceFromGUID(ByVal Guid As Guid) As Object
    'Set CreateInstanceFromGUID = VBA.GetObject(FilePath, ClassName)
End Function


Public Function Version() As Version
    Dim VM As Long
    'VM = API_get
    
    Set Version = New Version
    Call Version.Constructor2(VM)
End Function

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------
