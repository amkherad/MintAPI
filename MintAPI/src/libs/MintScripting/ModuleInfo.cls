VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ModuleInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim mName As String
Dim mDescription As String
Dim mVersion As Long
Dim mCompatibilityInfo(7) As Long
Dim mCopyright As String

Dim mDependenciesCount As Long
Dim mDependencies() As String

Dim mTranslationsCount As Long
Dim mTranslations() As String

Dim mCodeBehind As String

Public Function Create(ClassName As String) As IClass
    
End Function
