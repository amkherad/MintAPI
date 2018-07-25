VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProcessStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="ProcessStatic" static tags="">
''  <summary>
''    Provides process static methods.
''  </summary>
''  <using>Process</using>
''  <idea>Me , Qt , .net Framework , Java , Andriod activities.</idea>
''  <sample>
''      Process.Current. ... 'Returns current process.
''      stdout Process.OpenProcess("C:\testproc.exe").ProcessID
''  </sample>
''</class>

Option Base 0
Option Explicit

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "ProcessStatic" '#DO NOT REMOVE IT
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
Private Const TH32CS_SNAPPROCESS = &H2&


Private Declare Function API_GetCurrentProcess Lib "kernel32" Alias "GetCurrentProcess" () As Long
Private Declare Function API_GetCurrentProcessId Lib "kernel32" Alias "GetCurrentProcessId" () As Long
Private Declare Function API_CreateToolhelp32Snapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function API_Process32First Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, lppe As API_PROCESSENTRY32) As Long
Private Declare Function API_Process32Next Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, lppe As API_PROCESSENTRY32) As Long

Private Type API_PROCESSENTRY32
  dwSize As Long
  cntUsage As Long
  th32ProcessID As Long
  th32DefaultHeapID As Long
  th32ModuleID As Long
  cntThreads As Long
  th32ParentProcessID As Long
  pcPriClassBase As Long
  dwFlags As Long
  szExeFile As String * MAX_PATH
End Type

'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Variables>
'
Dim CurrentProcess_Static As Process
'
'</section>
'---------------------------------------------
'---------------------------------------------

'=============================================
'=============================================
'=============================================
'<section Creators>
'
Public Function Create() As Process
    
End Function
Public Function FromHandle(ByVal Handle As Long) As Process
    Dim cProc As Process
    Set cProc = New Process
    Call cProc.Constructor0(Handle)
    Set FromHandle = cProc
End Function
Public Function FromProcessID(ByVal ProcessID As Long) As Process
    Dim cProc As Process
    Set cProc = New Process
    Call cProc.Constructor1(ProcessID)
    Set FromProcessID = cProc
End Function

Public Function OpenProcess(ByVal Path As String, ByVal Arguments As String, Optional ByVal Environment As String) As Process
    
End Function
Public Function OpenProcessAs(ByVal Path As String, ByVal User As String, ByVal Arguments As String, Optional ByVal Environment As String) As Process
    
End Function
Public Function OpenPipedProcess(ByVal Path As String, ByVal Arguments As String, Optional ByVal Environment As String) As Process
    
End Function
Public Function OpenPipedProcessAs(ByVal Path As String, ByVal User As String, ByVal Arguments As String, Optional ByVal Environment As String) As Process
    
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

Public Property Get Current() As Process
    If CurrentProcess_Static Is Nothing Then
        Set CurrentProcess_Static = New Process
        Call CurrentProcess_Static.Constructor1(API_GetCurrentProcess)
    End If
    Set Current = CurrentProcess_Static
End Property


Public Function EndProcess(Optional ByVal ExitCode As Long = SUCCESSFULL) As Long
    Call Current.EndProcess(ExitCode)
End Function


Public Function CountProcesses() As Long
    
End Function
Public Function EnumProcesses() As Process()
    Dim lSnapShot As Long
    Dim lNextProcess As Long
    Dim tPE As API_PROCESSENTRY32
    Dim pList As New List
    'Dim p() As Process
    Dim pCount As Long

    ' Create snapshot
    lSnapShot = API_CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)
    If lSnapShot <> -1 Then
        ' Length of the structure
        tPE.dwSize = Len(tPE)
        ' Find first process
        lNextProcess = API_Process32First(lSnapShot, tPE)
        Do While lNextProcess
'            Found specified process
'            If sExeName = Left$(tPE.szExeFile, Len(sExeName)) And Len(sExeName) > 0 Then
'                Dim lProcess As Long
'                Dim lExitCode As Long
'                ' Open process
'                lProcess = OpenProcess(0, False, tPE.th32ProcessID)
'                ' Terminate process
'                TerminateProcess lProcess, lExitCode
'                ' Close handle
'                CloseHandle lProcess
                ' Add exe to list
                Dim cProc As Process
                Set cProc = New Process
                Call cProc.Constructor0(tPE.th32ProcessID)
            'End If
            ' Get next process
            lNextProcess = API_Process32Next(lSnapShot, tPE)
        Loop
        ' Close handle
        Call API_CloseHandle(lSnapShot)
    Else
        throw Exps.SystemCallFailureException
    End If
End Function

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------
