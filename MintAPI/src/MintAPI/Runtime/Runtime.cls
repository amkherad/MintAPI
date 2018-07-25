VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Runtime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="Runtime" static tags="">
''  <summary>
''    Provides some runtime operations.
''  </summary>
''  <using>Method</using>
''  <idea>Me</idea>
''  <sample>
''      Dim ASMCode() As Byte
''      ASMCode = myClass.CreateAssemblyCode()
''      Dim ManagedMethod As Method, retVal() As Byte
''      Set ManagedMethod = Runtime.CreateMethod("some_valid_name", Nothing, ByteArray(ASMCode))
''      Call ManagedMethod.Invoke(Nothing, 0) 'retVal does not used, because there is no return value.
''  </sample>
''</class>

Option Base 0
Option Explicit

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "Runtime" '#DO NOT REMOVE IT
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
Public Function CreateMethod(ByVal Name As String, _
        ByVal Prototype As MethodPrototype, _
        ByVal MethodBody As IMemoryBlock) As Method
    If Prototype Is Nothing Then throw Exps.ArgumentNullException("Prototype")
    If MethodBody Is Nothing Then throw Exps.ArgumentNullException("MethodBody")

    Dim VirtMem As IMemoryBlock, Size As Long
    Size = MethodBody.Limit
    If Size = 0 Then throw Exps.InvalidOperationException
    
    Set VirtMem = VirtualMemory.Allocate(Size, vmatCommit, mapFullControl)
    Call Memory.CopyBlock(VirtMem, MethodBody, Size)

    Dim Freezable As Freezable
    Set Freezable = Prototype
    If Not Freezable.Freeze Then throw Exps.AccessDeniedException

    Set CreateMethod = Method.FromReference(Name, VirtMem.Reference, Prototype)
    CreateMethod.IsOnTheFly = True
    CreateMethod.OnTheFlyLength = Size
    Call GC.AddToFamilyChain(CreateMethod, VirtMem)
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

'Public Function Compile(ByVal Compiler As String, Code As String) As Method
'
'End Function

'Public Function GetMethodBody(ByVal Method As Method) As ByteArray
'
'End Function

Public Function GetInterfaceVirtualTablePtr(ByVal Interface As Object) As Long
    GetInterfaceVirtualTablePtr = MemPtr(ObjPtr(Interface))
End Function
Public Function GetInterfaceVirtualTable(ByVal Interface As Object) As Vector '(Of Methods)
    
End Function

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------
