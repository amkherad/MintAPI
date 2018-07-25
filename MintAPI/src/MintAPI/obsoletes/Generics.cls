VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Generics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Base 0
Option Explicit
Const CLASSID As String = "Generics"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


'Implements IAliasable


Public Enum EqualToFlags
    AllMustEqual = 1
    SomeEqual = 2
    SomeNotEqual = 4
    AllMustNotEqual = 8

    SomeCompare = SomeEqual Or SomeNotEqual
    AllCompare = AllMustEqual Or AllMustNotEqual
    NotValue = 32
    AllValues = AllMustEqual Or SomeEqual Or SomeNotEqual Or AllMustNotEqual
End Enum


Private Declare Function API_ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Enum ShellWindowStyles
    swsHide = VbAppWinStyle.vbHide
    swsMaximizedFocus = VbAppWinStyle.vbMaximizedFocus
    swsMinimizedFocus = VbAppWinStyle.vbMinimizedFocus
    swsMinimizedNoFocus = VbAppWinStyle.vbMinimizedNoFocus
    swsNormalFocus = VbAppWinStyle.vbNormalFocus
    swsNormalNoFocus = VbAppWinStyle.vbNormalNoFocus
End Enum

Friend Sub Initialize()
    
End Sub


Public Function bnShiftLeft(bgNum As BigNumber, TimesToShift As Long) As BigNumber
    Set bnShiftLeft = bgNum.bnShiftLeft(TimesToShift)
End Function
Public Function bnShiftRight(bgNum As BigNumber, TimesToShift As Long) As BigNumber
    Set bnShiftRight = bgNum.bnShiftRight(TimesToShift)
End Function
Public Function bnRotateLeft(bgNum As BigNumber, TimesToRotate As Long) As BigNumber
    Set bnRotateLeft = bgNum.bnRotateLeft(TimesToRotate)
End Function
Public Function bnRotateRight(bgNum As BigNumber, TimesToRotate As Long) As BigNumber
    Set bnRotateRight = bgNum.bnRotateRight(TimesToRotate)
End Function

Public Function bnAdd(bgNum1 As BigNumber, bgNum2 As BigNumber) As BigNumber
    Set bnAdd = bgNum1.bnAdd(bgNum2)
End Function
Public Function bnSub(bgNum1 As BigNumber, bgNum2 As BigNumber) As BigNumber
    Set bnSub = bgNum1.bnSub(bgNum2)
End Function
Public Function bnMul(bgNum1 As BigNumber, bgNum2 As BigNumber) As BigNumber
    Set bnMul = bgNum1.bnMul(bgNum2)
End Function
Public Function bnDiv(bgNum1 As BigNumber, bgNum2 As BigNumber) As BigNumber
    Set bnDiv = bgNum1.bnDiv(bgNum2)
End Function
Public Function bnMod(bgNum1 As BigNumber, bgNum2 As BigNumber) As BigNumber
    Set bnMod = bgNum1.bnMod(bgNum2)
End Function
Public Function bnPow(bgNum1 As BigNumber, bgNum2 As BigNumber) As BigNumber
    Set bnPow = bgNum1.bnPow(bgNum2)
End Function
Public Function bnLog(bgNum As BigNumber, bgBaseNum As BigNumber) As BigNumber
    Set bnLog = bgNum.bnLog(bgBaseNum)
End Function



Public Function GlobalFilters(Optional Includes, Optional Excludes) As GlobalFilters
    Dim RetVal As GlobalFilters
    If IsMissing(Includes) Then
        Dim p(1) As String
        p(0) = "*": p(1) = "*.*"
        Includes = p
    End If
    If VarType(Includes) = (vbArray Or vbString) Then
        RetVal.IncludeTemplates = Includes
    Else
        ReDim RetVal.IncludeTemplates(0)
        RetVal.IncludeTemplates(0) = CStr(Includes)
    End If
    If Not IsMissing(Excludes) Then
        If VarType(Excludes) = (vbArray Or vbString) Then
            RetVal.ExcludeTemplates = Excludes
        Else
            ReDim RetVal.ExcludeTemplates(0)
            RetVal.ExcludeTemplates(0) = CStr(Excludes)
        End If
    End If
    GlobalFilters = RetVal
End Function
Public Function FileFilters(Optional Includes, Optional Excludes) As GlobalFilters
    Dim RetVal As GlobalFilters
    If IsMissing(Includes) Then
        Dim p(1) As String
        p(0) = "*": p(1) = "*.*"
        Includes = p
    End If
    If VarType(Includes) = (vbArray Or vbString) Then
        RetVal.IncludeTemplates = Includes
    Else
        ReDim RetVal.IncludeTemplates(0)
        RetVal.IncludeTemplates(0) = CStr(Includes)
    End If
    If Not IsMissing(Excludes) Then
        If VarType(Excludes) = (vbArray Or vbString) Then
            RetVal.ExcludeTemplates = Excludes
        Else
            ReDim RetVal.ExcludeTemplates(0)
            RetVal.ExcludeTemplates(0) = CStr(Excludes)
        End If
    End If
    FileFilters = RetVal
End Function


Private Function IncludeFilter_single(fs() As String, Expression As String) As Boolean
    Dim i As Long
    For i = 0 To ArraySize(fs) - 1
        If (Expression) Like CStr(fs(i)) Then
            IncludeFilter_single = True
            Exit Function
        End If
    Next
    IncludeFilter_single = False
End Function
Public Function IsFilterIncluded(Filter As GlobalFilters, Expression As String) As Boolean
    IsFilterIncluded = ((IncludeFilter_single(Filter.IncludeTemplates, Expression)) And (Not IncludeFilter_single(Filter.ExcludeTemplates, Expression)))
End Function

Public Sub InitializeRandomGenerator(): Call Randomize: End Sub
Public Function GetNextRandom(Optional UppderValue As Double = 1, Optional LowerValue As Double = 0, Optional Initialize_NewInstance_Of_RandomGenerator) As Double
    If Not IsMissing(Initialize_NewInstance_Of_RandomGenerator) Then _
        Call InitializeRandomGenerator
    GetNextRandom = LowerValue + ((UppderValue - LowerValue) * Rnd)
End Function

Public Function EqualTo(Flags As EqualToFlags, EnumerateValue, ParamArray Args() As Variant) As Boolean
    Dim iArgs() As Variant
    iArgs = Args
    EqualTo = baseMethods.EqualToArr(Flags, EnumerateValue, iArgs)
End Function

Public Function MemoryToByteArray(ByVal TargetAddress As Long, SourceSize As Long, Optional Times As Long = 1) As Byte()
    MemoryToByteArray = mint_var.MemoryToByteArray(TargetAddress, SourceSize, Times)
End Function
Public Sub CopyMemoryToByteArray(ByVal TargetAddress As Long, SourceSize As Long, targetByteArray() As Byte, Optional Times As Long = 1)
    Call mint_var.CopyMemoryToByteArray(TargetAddress, SourceSize, targetByteArray, Times)
End Sub
Public Sub ByteArrayToMemory(ByVal TargetAddress As Long, targetByteArray() As Byte, BytesToCopy As Long, Optional FillWithNull As Boolean = False)
    Call mint_var.ByteArrayToMemory(TargetAddress, targetByteArray, BytesToCopy, FillWithNull)
End Sub


Public Function RepeatString(Str As String, Count As Long, Optional Splitter As String = "") As String
    Dim i As Long
    For i = 1 To Count
        RepeatString = RepeatString & Str & Splitter
    Next
End Function



'Public Sub ByteArrayToParamArrayArgs(targetByteArray() As Byte, target)
'    Call baseMethods.ByteArrayToParamArrayArgs(targetByteArray, target)
'End Sub
Public Function BinaryCompare(A1, A2, Optional LengthToCompare As Long = -1) As Boolean
    BinaryCompare = baseMethods.BinaryCompare(A1, A2, LengthToCompare)
End Function
Public Function ArrayToByteArray(TargetArray) As Byte()
    ArrayToByteArray = baseMethods.ArrayToByteArray(TargetArray)
End Function
Public Function GetSubByteArray(targetByteArray() As Byte, ByVal StartIndex As Long, Optional ByVal Length As Long = -1) As Byte()
    GetSubByteArray = baseMethods.GetSubByteArray(targetByteArray, StartIndex, Length)
End Function
Public Function GetByteArraySomeLength(TargetArray() As Byte, Length As Long) As Byte()
    GetByteArraySomeLength = baseMethods.GetByteArraySomeLength(TargetArray, Length)
End Function
Public Function IsByteArrayNumeric(bArr() As Byte) As Boolean
    IsByteArrayNumeric = baseMethods.IsByteArrayNumeric(bArr)
End Function
Public Function IsByteArrayAlphabetic(bArr() As Byte) As Boolean
    IsByteArrayAlphabetic = baseMethods.IsByteArrayAlphabetic(bArr)
End Function
Public Function IsByteArrayLikeAnother(fA() As Byte, SA() As Byte) As Boolean
    IsByteArrayLikeAnother = baseMethods.IsByteArrayLikeAnother(fA, SA)
End Function
Public Function ByteArrayToUpper(bArr() As Byte) As Byte()
    ByteArrayToUpper = baseMethods.ByteArrayToUpper(bArr)
End Function
Public Function ByteArrayToLower(bArr() As Byte) As Byte()
    ByteArrayToLower = baseMethods.ByteArrayToLower(bArr)
End Function
Public Function FillArray(TargetArray, whattofill, OutputSize As Long, Optional FillToLeft = True)
    FillArray = baseMethods.FillArray(TargetArray, whattofill, OutputSize, FillToLeft)
End Function
Public Sub InsertArrayIndex(TargetArray, Index, Item)
    Call baseMethods.InsertArrayIndex(TargetArray, Index, Item)
End Sub
Public Sub InsertArrayIndexArray(TargetArray, Index, insertArray, Optional Length As Long = -1)
    Call baseMethods.InsertArrayIndexArray(TargetArray, Index, insertArray, Length)
End Sub
Public Function AppendToArray(sourceArray, Item)
    Call baseMethods.AppendToArray(sourceArray, Item)
End Function
Public Function AppendArrayToArray(sourceArray, TargetArray)
    Call baseMethods.AppendArrayToArray(sourceArray, TargetArray)
End Function


Public Function SetStreamState(State As Boolean, inoutState As Boolean, State_LOCK As String, Optional Reserved)
    Call pubs.mint_setstream_state(State, inoutState, State_LOCK, Reserved)
End Function


Public Function ShellExecute(ByVal hWnd As Long, ByVal Operation As String, ByVal FileName As String, ByVal Parameters As String, ByVal Directory, Optional ByVal ShowCommand As ShellWindowStyles = ShellWindowStyles.swsNormalFocus) As Long
    ShellExecute = API_ShellExecute(hWnd, Operation, FileName, Parameters, Directory, ShowCommand)
End Function
