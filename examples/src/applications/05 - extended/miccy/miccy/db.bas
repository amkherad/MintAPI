Attribute VB_Name = "db"
Option Explicit

Public Const CONFIGVALIDATION As Long = 824535405
Public Const CONFIGVALIDATION2 As Long = 2036556581
Public Const CONFIGLASTVALIDATION1 As Long = 1667459437
Public Const CONFIGLASTVALIDATION2 As Long = 828662137

Private Const SPECNEXTISRECORD As Long = &H1
Private Const SPECNEXTISINFO As Long = &H10

Public Const TBL_VALIDATION As Long = 1
Public Const TBL_LICENCEKEY As Long = TBL_VALIDATION + 4
Public Const TBL_MULTIINSTANCE As Long = TBL_LICENCEKEY + 32
Public Const TBL_BUFFERDATAMODE As Long = TBL_MULTIINSTANCE + 2
Public Const TBL_BUFFERDATAREMEMBER As Long = TBL_BUFFERDATAMODE + 4
Public Const TBL_LOADALLPLUGINSINSTARTUP As Long = TBL_BUFFERDATAREMEMBER + 2
Public Const TBL_ALLOWPLUGINSTOCHANGEGLOBALSETTINGS As Long = TBL_LOADALLPLUGINSINSTARTUP + 2
Public Const TBL_ALLOWPLUGINSTOINSTALLORUNINSTALLOTHERS As Long = TBL_ALLOWPLUGINSTOCHANGEGLOBALSETTINGS + 2
Public Const TBL_ALLOWPLUGINSTOTERMINATEPROCCESS As Long = TBL_ALLOWPLUGINSTOINSTALLORUNINSTALLOTHERS + 2
Public Const TBL_ALLOWPLUGINSTOCONNECTTONETWORK As Long = TBL_ALLOWPLUGINSTOTERMINATEPROCCESS + 2
Public Const TBL_AUTOCHECKFORUPDATES As Long = TBL_ALLOWPLUGINSTOCONNECTTONETWORK + 2
Public Const TBL_SHOWAPPLICATIONTIPSMODE As Long = TBL_AUTOCHECKFORUPDATES + 4
Public Const TBL_OS As Long = TBL_SHOWAPPLICATIONTIPSMODE + 4

Public Const TBL_PLUGINSCOUNT As Long = TBL_OS + 2
Public Const TBL_FILTRSCOUNT As Long = TBL_PLUGINSCOUNT + 4
Public Const TBL_ACTIONSCOUNT As Long = TBL_FILTRSCOUNT + 4
Public Const TBL_GLOBALSCOUNT As Long = TBL_ACTIONSCOUNT + 4
Public Const TBL_PLUGINSSTART As Long = TBL_GLOBALSCOUNT + 4
Public Const TBL_FILTRSSTART As Long = TBL_PLUGINSSTART + 4
Public Const TBL_ACTIONSSTART As Long = TBL_FILTRSSTART + 4
Public Const TBL_GLOBALSSTART As Long = TBL_ACTIONSSTART + 4
Public Const TBL_OVERALLACTIONSAFTERLASTFRAGMENTATION As Long = TBL_GLOBALSSTART + 4
Public Const TBL_VALIDATION2 As Long = TBL_OVERALLACTIONSAFTERLASTFRAGMENTATION + 4
Public Const TBL_HISTORYCOUNT As Long = TBL_VALIDATION2 + 4
Public Const TBL_DYNDATASTART As Long = TBL_HISTORYCOUNT + 4

Public Const RAPLUGINS As Long = TBL_PLUGINSCOUNT
Public Const RAFILTERS As Long = TBL_FILTRSCOUNT
Public Const RAACTIONS As Long = TBL_ACTIONSCOUNT
Public Const RAGLOBALS As Long = TBL_GLOBALSCOUNT


Public Type GLBTP
    Prev As Long
    Next As Long
    uniqueID As String * 50
    Name As String * 20
    Source As String * 50
    entryName As String * 20
    Description As String * 256
End Type

Public Type GlobalProperties
    MultiInstance  As Boolean
    BufferDataMode As Long
    BufferDataRemember As Boolean
    LoadAllPluginsInStartup As Boolean
    AllowPluginsToChangeGlobalSettings As Boolean
    AllowPluginsToInstallOrUninstallOthers As Boolean
    AllowPluginsToTerminateProccess As Boolean
    AllowPluginsToConnectToNetwork As Boolean
    AutoCheckForUpdates As Long
    ShowApplicationTipsMode As Long
    CountHistoryItems As Long
    DynamicStart As Long
End Type

Public plugins() As GLBTP
Public pluginsCount As Long
Public pluginsEdited As Boolean

Public filters() As GLBTP
Public filtersCount As Long
Public filtersEdited As Boolean

Public actions() As GLBTP
Public actionsCount As Long
Public actionsEdited As Boolean

Public globals() As GLBTP
Public globalsCount As Long
Public globalsEdited As Boolean

Public gp As GlobalProperties
Public gpMustSave As Boolean

Dim configFL As Long
Public cFL As Long
Dim configWritable As Boolean

Public Sub ReadConfig()
    Call OpenConfig(True)
    Debug.Print "Configuration Reading"
    Get #configFL, TBL_MULTIINSTANCE, gp.MultiInstance
    Get #configFL, TBL_BUFFERDATAMODE, gp.BufferDataMode
    Get #configFL, TBL_BUFFERDATAREMEMBER, gp.BufferDataRemember
    Get #configFL, TBL_LOADALLPLUGINSINSTARTUP, gp.LoadAllPluginsInStartup
    Get #configFL, TBL_ALLOWPLUGINSTOCHANGEGLOBALSETTINGS, gp.AllowPluginsToChangeGlobalSettings
    Get #configFL, TBL_ALLOWPLUGINSTOINSTALLORUNINSTALLOTHERS, gp.AllowPluginsToInstallOrUninstallOthers
    Get #configFL, TBL_ALLOWPLUGINSTOTERMINATEPROCCESS, gp.AllowPluginsToTerminateProccess
    Get #configFL, TBL_ALLOWPLUGINSTOTERMINATEPROCCESS, gp.AllowPluginsToTerminateProccess
    Get #configFL, TBL_ALLOWPLUGINSTOCONNECTTONETWORK, gp.AllowPluginsToConnectToNetwork
    Get #configFL, TBL_AUTOCHECKFORUPDATES, gp.AutoCheckForUpdates
    Get #configFL, TBL_SHOWAPPLICATIONTIPSMODE, gp.ShowApplicationTipsMode
    Get #configFL, TBL_HISTORYCOUNT, gp.CountHistoryItems
    Debug.Print "Configuration Readed"
End Sub
Public Sub OpenConfig(Optional isValidFile As Boolean = True)
    If configFL = 0 Then
        Dim validLong As Long
        configFL = FreeFile
        cFL = configFL
        configWritable = True
        Open App.Path & "\config" For Binary As #configFL
        clog "db: Opening File: " & App.Path & "\config"
        If isValidFile Then
            Get #configFL, TBL_VALIDATION, validLong
            Debug.Print "db: Validation 1 Read:" & validLong & " From Column " & TBL_VALIDATION & " Must Equals To :" & CONFIGVALIDATION
            If validLong <> CONFIGVALIDATION Then throw Exps.InvalidFileException
            Get #configFL, TBL_VALIDATION2, validLong
            Debug.Print "db: Validation 2 Read:" & validLong & " From Column " & TBL_VALIDATION2 & " Must Equals To :" & CONFIGVALIDATION2
            If validLong <> CONFIGVALIDATION2 Then throw Exps.InvalidFileException
            
            Get #configFL, LOF(configFL) - 7, validLong
            Debug.Print "db: Validation 3 Read:" & validLong & " From Column " & LOF(configFL) - 7 & " Must Equals To :" & CONFIGLASTVALIDATION1
            If validLong <> CONFIGLASTVALIDATION1 Then throw Exps.InvalidFileException
            Get #configFL, LOF(configFL) - 3, validLong
            Debug.Print "db: Validation 4 Read:" & validLong & " From Column " & LOF(configFL) - 3 & " Must Equals To :" & CONFIGLASTVALIDATION2
            If validLong <> CONFIGLASTVALIDATION2 Then throw Exps.InvalidFileException
            configWritable = False
        End If
    End If
End Sub
Public Sub SaveConfig()
    
End Sub
Private Sub SaveConfig_p(ByVal Path As String)
    Debug.Print "Configuration Saving"
    Dim fl As Long
    fl = FreeFile
    Open Path For Binary As #fl
    
    Put #fl, TBL_VALIDATION, CONFIGVALIDATION
    Put #fl, TBL_VALIDATION2, CONFIGVALIDATION2
            
    Put #fl, TBL_MULTIINSTANCE, gp.MultiInstance
    Put #fl, TBL_BUFFERDATAMODE, gp.BufferDataMode
    Put #fl, TBL_BUFFERDATAREMEMBER, gp.BufferDataRemember
    Put #fl, TBL_LOADALLPLUGINSINSTARTUP, gp.LoadAllPluginsInStartup
    Put #fl, TBL_ALLOWPLUGINSTOCHANGEGLOBALSETTINGS, gp.AllowPluginsToChangeGlobalSettings
    Put #fl, TBL_ALLOWPLUGINSTOINSTALLORUNINSTALLOTHERS, gp.AllowPluginsToInstallOrUninstallOthers
    Put #fl, TBL_ALLOWPLUGINSTOTERMINATEPROCCESS, gp.AllowPluginsToTerminateProccess
    Put #fl, TBL_ALLOWPLUGINSTOTERMINATEPROCCESS, gp.AllowPluginsToTerminateProccess
    Put #fl, TBL_ALLOWPLUGINSTOCONNECTTONETWORK, gp.AllowPluginsToConnectToNetwork
    Put #fl, TBL_AUTOCHECKFORUPDATES, gp.AutoCheckForUpdates
    Put #fl, TBL_SHOWAPPLICATIONTIPSMODE, gp.ShowApplicationTipsMode
    Put #fl, TBL_HISTORYCOUNT, gp.CountHistoryItems
    
    Put #fl, LOF(fl) + 1, CONFIGLASTVALIDATION1
    Put #fl, LOF(fl) + 1, CONFIGLASTVALIDATION2
    
    Close #fl
    Debug.Print "Configuration Saved"
End Sub
Public Sub FlushConfig()
    Call CloseConfig
    Call OpenConfig
End Sub
Public Sub CloseConfig()
    If configFL <> 0 Then
        Call EndFileBlocks
        Call SaveConfig
        Close #configFL
        configFL = 0
        cFL = 0
    End If
End Sub
Public Function isConfig() As Boolean
    isConfig = (configFL <> 0)
End Function
Public Sub PrepareForWrite()
    If configWritable Then Exit Sub
    configWritable = True
End Sub
Public Sub EndFileBlocks()
    Dim validLong1 As Long, validLong2 As Long
    Get #configFL, LOF(configFL) - 7, validLong1
    Get #configFL, LOF(configFL) - 3, validLong2
    If Not ((validLong1 = CONFIGLASTVALIDATION1) And (validLong2 = CONFIGLASTVALIDATION2)) Then
        Put #configFL, LOF(configFL) + 1, CONFIGLASTVALIDATION1
        Put #configFL, LOF(configFL) + 1, CONFIGLASTVALIDATION2
    End If
End Sub

Public Function installPlugin(uniqueID As String, Name As String, Source As String, entryName As String, Description As String)
    
End Function
Public Sub uninstallPlugin()
    
End Sub
Public Function InstallAction()
    
End Function
Public Sub UninstallAction()
    
End Sub
Public Function InstallFilter()
    
End Function
Public Sub UninstallFilter()
    
End Sub
