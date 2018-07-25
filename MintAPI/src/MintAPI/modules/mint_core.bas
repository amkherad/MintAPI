Attribute VB_Name = "mint_core"
Option Explicit

Public Const MINTCORE_DLL_FILENAME As String = "mintcore.dll"

Private mint_core_handle As Long
Private mint_core_path As String

Public Mint_Core_NotFound As Boolean

Private Declare Function MC_InitializeMintCore Lib "MintCore" Alias "InitializeMintCore" () As Long

Public Sub Construct()
    'If Not API_Register_MintAPI_Lib Then throw Exps.InvalidCallException
    mint_core_path = Path.Combine(App.Path, MINTCORE_DLL_FILENAME)
    If Not File.Exists(mint_core_path) Then _
        mint_core_path = FindMintCoreAnyway
    
    If mint_core_path = "" Then _
        mint_core_path = TryToInstallMintCore
    
    If mint_core_path = "" Then
        Mint_Core_NotFound = True
    Else
        Mint_Core_NotFound = False
        mint_core_handle = API_LoadLibrary(mint_core_path)
        Call MC_InitializeMintCore
    End If
End Sub

Private Function FindMintCoreAnyway() As String
'    Exit Function
'    Dim R As Registry, dirPath As String, DllPath As String
'    Set R = Registry(APP_REGISTRYPATH).cd("MintCore")
'    dirPath = R.GetValue("dir").ToString
'    DllPath = R.GetValue("dll_path").ToString
'
'    If File(DllPath).Exists Then
'        FindMintCoreAnyway = DllPath
'        Exit Function
'    End If
'
'    Dim fList() As String
'    fList = Directory(dirPath).FilteredFileNames(FileFilters(StringArray("mintcore*.dll"), ""))
'    'mintcore[0].dll
'    Dim i As Long, mxStr As String
'    Dim maxI As Long, Max As Long, cVal As Long
'    maxI = -1
'    For i = 0 To ArraySize(fList) - 1
'        If LCase$(Left(fList(i), 8)) = "mintcore" Then
'            If LCase$(Right(fList(i), 3)) = "dll" Then
'                mxStr = Mid(fList(i), 9, 1)
'                If mxStr <> "." Then
'                    cVal = Val(mxStr)
'                    If cVal > Max Then
'                        maxI = i
'                        Max = cVal
'                    End If
'                End If
'            End If
'        End If
'    Next
'
'    If maxI >= 0 Then _
'        FindMintCoreAnyway = ConcatPath(dirPath, fList(maxI))
End Function


Private Function TryToInstallMintCore() As String
    
End Function
