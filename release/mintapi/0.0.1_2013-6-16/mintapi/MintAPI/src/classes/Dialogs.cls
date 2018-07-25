VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Dialogs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

Option Base 0
Option Explicit
Const CLASSID As String = "Dialogs"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


Implements IAliasable


Friend Sub Initialize()

End Sub

Public Function Show()

End Function

Public Sub ShowTooltip()

End Sub

Public Function GetDouble() As Double

End Function
Public Function GetInt() As Integer

End Function
Public Function GetLong() As Long

End Function
Public Function GetEnumerate() As Long

End Function
Public Function GetString() As String

End Function
Public Function GetRegistryPath() As Registry

End Function

Public Function ChooseDirectory(Optional ByVal OwnerForm As Long = 0, Optional Title As String = "", _
            Optional ByVal RootDir As SpecialFolders = SpecialFolders.sfALL, Optional ByVal CustomRootDir As Directory = Nothing, _
            Optional ByVal StartDir As Directory = Nothing, Optional NewStyle As Boolean = True, _
            Optional IncludeFiles As Boolean = False, Optional ShowNewFolderButton As Boolean = True, _
            Optional OkButtonText As String = "") As Directory
    Dim s As String
    Dim custom_Dir As String, start_Dir As String
    If Not CustomRootDir Is Nothing Then custom_Dir = CustomRootDir.AbsolutePath
    If Not StartDir Is Nothing Then start_Dir = StartDir.AbsolutePath
    s = Dialogs_Browse(OwnerForm, Title, RootDir, custom_Dir, start_Dir, NewStyle, IncludeFiles, ShowNewFolderButton, OkButtonText)
    Dim d As New Directory
    Call d.Initialize(s)
    Set ChooseDirectory = d
End Function
Public Function ChooseDirectories() As Directory()

End Function
Public Function ChooseFile() As File

End Function
Public Function ChooseFiles() As File()

End Function
Public Function ChooseFont() As StdFont

End Function
Public Function ChooseColor() As Colors

End Function
Public Function MessageBox(Optional ByVal OwnerForm As Long = 0, Optional Title As String, Optional Message As String, _
            Optional Buttons As MessageBoxButtons = MessageBoxButtons.mbbOKOnly, _
            Optional DefaultButton As MessageBoxDefaultButton = MessageBoxDefaultButton.mbdButton1, _
            Optional Icon As MessageBoxIcons = MessageBoxIcons.siNoIcon, _
            Optional Styles As MessageBoxStyles = MessageBoxStyles.mbsNoSpecifiedStyle, _
            Optional Button1Text As String, Optional Button2Text As String, Optional Button3Text As String, Optional Button4Text As String, _
            Optional ByVal OnOKAction As Action = Nothing, Optional ByVal OnCancelAction As Action = Nothing) As DialogResults

End Function



Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property
