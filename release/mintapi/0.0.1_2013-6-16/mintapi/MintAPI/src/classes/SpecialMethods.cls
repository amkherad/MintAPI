VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SpecialMethods"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Base 0
Option Explicit
Const CLASSID As String = "SpecialMethods"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


Implements IAliasable


Public Enum ProcessorFeatures
    PF_FLOATING_POINT_PRECISION_ERRATA = 0
    PF_FLOATING_POINT_EMULATED = 1
    PF_COMPARE_EXCHANGE_DOUBLE = 2
    PF_MMX_INSTRUCTIONS_AVAILABLE = 3
    PF_XMMI_INSTRUCTIONS_AVAILABLE = 6
    PF_3DNOW_INSTRUCTIONS_AVAILABLE = 7
    PF_RDTSC_INSTRUCTION_AVAILABLE = 8
    PF_PAE_ENABLED = 9
End Enum

#If New7API Then
    Private Declare Function API_IsProcessorFeaturePresent Lib "libkernel0.MintAPI.dll" _
        Alias "lk_IsProcessorFeaturePresent" (ByVal ProcessorFeature As Long) As Long
#Else
    Private Declare Function API_IsProcessorFeaturePresent Lib "kernel32" _
        Alias "IsProcessorFeaturePresent" (ByVal ProcessorFeature As Long) As Long
#End If

Friend Sub Initialize()

End Sub
Public Function RegisterLibraryLicense(LICENSE As String, Info As String) As Boolean
    RegisterLibraryLicense = licensing.RegisterLibraryLicense(LICENSE, Info)
End Function


Public Sub setExceptionState(Object As Object, Optional ByVal State As Boolean = True)

End Sub
Public Sub RestartExceptionStates(Object As Object)

End Sub

Public Sub RegisterProvider(ProviderEngine As String, Alias As String, Provider As IProvider)
    Call mint_type_provider.RegisterProvider(ProviderEngine, Alias, Provider)
End Sub
Public Sub ChangeProvider(ProviderEngine As String, Alias As String, Provider As IProvider)
    Call mint_type_provider.ChangeProvider(ProviderEngine, Alias, Provider)
End Sub
Public Sub UnRegisterProvider(ProviderEngine As String, Alias As String)
    Call mint_type_provider.UnRegisterProvider(ProviderEngine, Alias)
End Sub
Public Function Provide(ProviderEngine As String, Alias As String, Args As ArgumentList, ByRef retVal) As Boolean
    Provide = mint_type_provider.Provide(ProviderEngine, Alias, Args, retVal)
End Function


Public Function fgetf(hndl As Long) As File

End Function
Public Function fid(objFile As File) As Long

End Function
Public Function fopen(strPath As String) As Long

End Function
Public Sub fclose(hndl As Long)

End Sub

Public Function LastChooseFileReadOnlyState() As Boolean
    LastChooseFileReadOnlyState = pubs.mint_api_dialogs_last_choose_file_read_only_flag_state
End Function


Public Function GetObject(ByVal ObjectAddress As Long) As Object
    Dim X As Object
    Call API_CopyMemory(X, ObjectAddress, 4)
    Set GetObject = X
End Function

Public Function CreateLanguage(Name As String, Key As String, Region As String, Description As String, RightToLeft As Boolean) As Language
    Dim l As New Language
    Call l.SetAttributes(Name, Key, Region, Description, RightToLeft)
    Set CreateLanguage = l
End Function
Public Function CountLanguageRecords(ByVal lang As Language) As Long
    If lang Is Nothing Then throw ArgumentNullException
    CountLanguageRecords = lang.Count
End Function
Public Function CheckLanguageRecordExists(ByVal lang As Language, Key As String) As Boolean
    If lang Is Nothing Then throw ArgumentNullException
    CheckLanguageRecordExists = lang.translationExists(Key)
End Function
Public Function LanguageRecordKey(ByVal lang As Language, RecordIndex As Long) As String
    If lang Is Nothing Then throw ArgumentNullException
    LanguageRecordKey = lang.translationKey(RecordIndex)
End Function
Public Function LanguageRecordTranslation(ByVal lang As Language, RecordIndex As Long) As String
    If lang Is Nothing Then throw ArgumentNullException
    LanguageRecordTranslation = lang.translationValue(RecordIndex)
End Function
Public Sub AppendLanguageRecord(ByVal lang As Language, Key As String, Tranlation As String)
    If lang Is Nothing Then throw ArgumentNullException
    Call lang.appendTranslate(Key, Tranlation)
End Sub
Public Sub ChangeLanguageRecord(ByVal lang As Language, Key As String, NewTranlation As String)
    If lang Is Nothing Then throw ArgumentNullException
    Call lang.changeTranslate(Key, NewTranlation)
End Sub
Public Sub RemoveLanguageRecord(ByVal lang As Language, Key As String)
    If lang Is Nothing Then throw ArgumentNullException
    If Not lang.removeTranslate(Key) Then throw ItemNotExistsException
End Sub
Public Sub RemoveLanguageAllRecords(ByVal lang As Language)
    If lang Is Nothing Then throw ArgumentNullException
    Call lang.removeAll
End Sub
Public Sub CopyLanguageRecordsFromAnother(LanguageDestination As Language, LanguageSource As Language)
    If LanguageDestination Is Nothing Or LanguageSource Is Nothing Then throw ArgumentNullException
    Call LanguageDestination.CopyRecordsFrom(LanguageSource)
End Sub
Public Sub CopyLanguageFromAnother(LanguageDestination As Language, LanguageSource As Language)
    If LanguageDestination Is Nothing Or LanguageSource Is Nothing Then throw ArgumentNullException
    Call LanguageDestination.CopyFrom(LanguageSource)
End Sub
Public Sub AppendLanguageRecordsToAnother(LanguageDestination As Language, LanguageSource As Language)
    If LanguageDestination Is Nothing Or LanguageSource Is Nothing Then throw ArgumentNullException
    Call LanguageSource.appendTo(LanguageDestination)
End Sub

Public Function CreateFileSectionStream() As ITargetStream
    
End Function

Public Sub OpenURL(ByVal hWnd As Long, ByVal URL As String)
    Call shellMethods.OpenURL(hWnd, URL)
End Sub

Public Function IniRead(Path As String, Section As String, KeyName As String, Optional Default As String) As Variant
    IniRead = ReadINIFile(Path, Section, KeyName, Default)
End Function
Public Function IniWrite(Path As String, Section As String, KeyName As String, KeyValue As String) As Boolean
    IniWrite = WriteINIFile(Path, Section, KeyName, KeyValue)
End Function

Public Sub PermanentDeleteFile(Path As String)
    
End Sub
Public Function IsProcessorFeaturePresent(ProcessorFeatures As ProcessorFeatures) As Boolean
    IsProcessorFeaturePresent = (API_IsProcessorFeaturePresent(ProcessorFeatures) <> 0)
End Function

Public Function RegisterActiveX(Path As String) As Long
    RegisterActiveX = pubs.RegisterActiveX(Path)
End Function
Public Function UnRegisterActiveX(Path As String) As Long
    UnRegisterActiveX = pubs.UnRegisterActiveX(Path)
End Function

Public Function Copyright() As String
    Copyright = Mtr("Copyright (C) 2012 by Ali Mousavi Kherad alimousavikherad@gmail.com")
End Function
Public Sub AboutMintAPI(Optional ByVal Modal As Boolean = False)
    Call modMain.AboutMintAPI(Modal)
End Sub

Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property
