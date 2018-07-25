Attribute VB_Name = "modMain"
'@PROJECT_LICENSE

'by Ali Mousavi Kherad (GPL-v2, LGPL-v3)
'Free to use and distribute only if including my name and email as (alimousavikherad@gmail.com)!
'-----------------------------------------------------------------------
'|                          About the author                           |
'-----------------------------------------------------------------------
'| I have 20 years old.
'| And I am a system programmer,
'| Studying Computer Software Engineering
'| At Govenrnment Technical
'| Shamsipour University in Tehran
'|
'| My Projects:
'| 01. yLib:      2D CPU/GPU Game Engine
'|                (when I was begining programmer)
'| 02. MintAPI    (you'r seeing)
'| 03. doubleAC:  Advanced dynamic/scriptable
'|                accounting system. (commercial, Qt based)
'| 04. RoNOTE:    (planning) Advanced audio library and program.
'|                Acctually RoNOTE stand to establish all
'|                Guitar Pro's features to provide an open-source
'|                And free software for musicians.
'|                The main software is an note (music notes) editor,
'|                And also it provides features like Cubase, FLStudio
'|                For editor to be free of using mixer and composer
'|                Applications.
'|                (designing on Qt), Cross-Platform, (device supported)
'|
'|  And I made all these by magic of Heavy Metal !
'-----------------------------------------------------------------------

Option Explicit
Option Base 0
Const CLASSID As String = "modMain"

Public Const MINTAPI_FILECODE_Mint As Long = 1953392973 ' Mint

Public Const APPLICATIONDOMAIN As String = "com.mintapi"
Public Const APPLICATIONID As String = "MintAPI"
Public Const APP_VERSIONTAGS As String = "blackbox"
Public Const APP_VERSIONSTRING As String = "0.0.1.2015 " & APP_VERSIONTAGS
Public Const APP_RELEASEYEAR As Long = 2015
Public Const APP_WEBSITE As String = "mintapi.com"
Public Const APP_SERVICEWEBSITE As String = "soap." & APP_WEBSITE
Public Const APP_HELPLINK As String = APP_WEBSITE & "/help"
Public Const APP_UPDATELINK As String = APP_WEBSITE & "/update"
Public Const APP_SUPPORTLINK As String = APP_WEBSITE & "/support"

Public Const APP_GUID As String = ""

Public Const APP_PRODUCTCODE As String = "mintapi0000012015greenleafpxAB" ' 30 chars
Public Const APP_PRODUCTCODE50 As String = APP_PRODUCTCODE & "22xxxxoToloboxxxxxxx"

Public Const APP_REGISTRYPATH As String = "HKEY_LOCAL_MACHINE\SOFTWARE\MintAPI"
Public Const APP_REGISTRYPATH_USER As String = "HKEY_CURRENT_USER\SOFTWARE\MintAPI"

'Private Const API_ICC_USEREX_CLASSES = &H200

'Private Type API_tagInitCommonControlsEx
'   lngSize As Long
'   lngICC As Long
'End Type
'Private Declare Sub API_InitCommonControls Lib "comctl32" Alias "InitCommonControls" ()
'Private Declare Function API_InitCommonControlsEx Lib "comctl32" Alias "InitCommonControlsEx" (Iccex As API_tagInitCommonControlsEx) As Boolean


'New7API : determines compiling in win7 environment and remove kernel32 dll.

'=================================================================================================
'=================================================================================================
'=================================================================================================

Public Sub Main()
    'Call mint_core.Construct           'Obsolete
    Call mint_application.Construct     'IMPORTANT! Get Debugger State.
    Call mint_assemblies.Construct      'IMPORTANT! Initializes The Assemblies And MintHelper.
    'Call InitializeCommonControls      'Does not work in link library.'Obsolete
    
    'Call mint_gvariables.mint_register_safethread_globalvariable("license", 0)'Obsolete

    'Call kernelMethods.Initialize      'Obsolete
    'Call baseMethods.Initialize        'Obsolete
    'Call baseMethods2.Initialize       'Obsolete
    Call bitOperations.Initialize       'Obsolete IMPORTANT! Initializes the bit operators.
    'Call uiMethods.Initialize          'Obsolete
    'Call gdiMethods.Initialize         'Obsolete
    'Call shellMethods.Initialize       'Obsolete
    
    'Call mint_config.Construct         'Important to Initialize the config manager system.
    
    'Call CheckIfNotInstalled            'Installs the MintAPI environment.
    'Call DllLoadConfiguration          'Loads basic Dll configuration. (This method used to optimize the runtime execution)
End Sub
''<summary>Initializes the common controls for the process.</summary>
'Private Sub InitializeCommonControls() 'Does not work in link library.
'    Dim Iccex As API_tagInitCommonControlsEx
'    With Iccex
'        .lngSize = Len(Iccex)
'        .lngICC = API_ICC_USEREX_CLASSES
'    End With
'    Call API_InitCommonControlsEx(Iccex)
'End Sub

''<summary>Translates a MintAPI standard text.</summary>
''<retval>A translated string of the specified Key.</retval>
Public Function Mtr(Key As String) As String
    Mtr = Key
End Function

'Public Sub CheckIfNotInstalled()
'
'End Sub

'===========================================
''[<function></function>]
''<summary></summary>
'---
''<exceptions><exception></exception></exceptions>
''<events></events>
''<signals></signals>
''<params><param name="" type="" default="" in/out byval/byref onnull=""></param></params>
''<retval></retval>
''<sample></sample>
''<remarks></remarks>
''<see></see>
''<deprecated/>
''<obsolete/>
''<version 5 />
''<cargs count=""> <carg type=""></carg> </cargs> //ConstructorArg()
'--------------------
''<class interface constructor static></class>
''<using></using>
''<summary></summary>
''<sample></sample>
''<idea></idea>
''<remarks></remarks>
''<constructors><constructor></constructor></constructors>

'=============================================
'=============================================
'=============================================
'<section >
'
'Meta Data   =   Signals/Slots/Properties/Notifications/Defaults
'API Declarations
'Variables
'Constructors
'Body Members
'IObject Implementations
'ICloneable Implementations
'Private Helpers
'Class Events
'
'</section>
'---------------------------------------------
'---------------------------------------------
'---------------------------------------------

'Attribute asdf.VB_Description = "Ali"
'Attribute asdf.VB_UserMemId = 0'DEFAULT MEMBER
'Attribute asdf.VB_HelpID = 10
'BEGIN
'  MultiUse = -1  'True
'  Persistable = 0  'NotPersistable
'  DataBindingBehavior = 0  'vbNone
'  DataSourceBehavior = 0   'vbNone
'  MTSTransactionMode = 0   'NotAnMTSObject
'End
'Attribute VB_Name = "tApplication"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

'*FCOMPLETE*




'DISPID_COLLECT -8
'The Collect property. You use this property if the method you are calling through Invoke is an accessor function.
'
'DISPID_CONSTRUCTOR -6
'The C++ constructor function for the object.
'
'DISPID_DESTRUCTOR -7
'The C++ destructor function for the object.
'
'DISPID_EVALUATE -5
'The Evaluate method. This method is implicitly invoked when the ActiveX client encloses the arguments in square brackets. For example, the following two lines are equivalent:
'X.[A1:C1].Value = 10
'X.Evaluate("A1:C1").Value = 10
'
'DISPID_NEWENUM -4
'The _NewEnum property. This special, restricted property is required for collection objects. It returns an enumerator object that supports IEnumVARIANT, and should have the restricted attribute specified.
'
'DISPID_PROPERTYPUT -3
'The parameter that receives the value of an assignment in a PROPERTYPUT.
'
'DISPID_UNKNOWN -1
'The value returned by IDispatch::GetIDsOfNames to indicate that a member or parameter name was not found.
'
'DISPID_VALUE 0
'The default member for the object. This property or method is invoked when an ActiveX client specifies the object name without a property or method.
