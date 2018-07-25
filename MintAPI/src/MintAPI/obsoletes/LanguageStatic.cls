VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LanguageStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

Option Base 0
Option Explicit
Const CLASSID As String = "Language"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING

''<class static>
''  <summary>Provides static methods for <strong>Language</strong> class.</summary>
''  <see>Language</see>
''  <idea>Me , Qt</idea>
''</class>

'=============================================
'=============================================
'=============================================
'<section Static Creators>
'
Public Function Create(ByVal Name As String, ByVal Key As String, ByVal Region As String, ByVal Description As String, ByVal RightToLeft As Boolean, Optional ByVal Encoding As Object) As Language
    Dim Lang As New Language
    Call Lang.Constructor0(Name, Key, Region, Description, RightToLeft, Encoding)
    Set Create = Lang
End Function
Public Function FromFile(ByVal Path As String) As Language
    Dim l As New Language
    Call l.FromFile(Path)
    Set FromFile = l
End Function
'
'</section>
'---------------------------------------------
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Static Methods>
'

'
'</section>
'---------------------------------------------
'---------------------------------------------
'---------------------------------------------
