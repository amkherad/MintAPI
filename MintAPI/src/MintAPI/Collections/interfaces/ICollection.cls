VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

Option Base 0
Option Explicit
Const CLASSID As String = "ICollection"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


Public Function GetEnumerator() As IEnumerator: End Function
Public Function NewEnum() As IUnknown: End Function
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Public Property Get Count() As Long: End Property
Public Sub CopyTo(ByRef Arr As Variant, ByVal Index As Long): End Sub

