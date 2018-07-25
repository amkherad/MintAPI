VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NumberStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

''<class name="NumberStatic" static tags="">
''  <summary>
''    Some class description.
''  </summary>
''  <using>Some class alias.</using>
''  <idea>Me</idea>
''  <sample></sample>
''</class>

Option Base 0
Option Explicit

'=============================================
'=============================================
'=============================================
'<section Meta Data>
'
Const CLASSID As String = "NumberStatic" '#DO NOT REMOVE IT
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
Public Property Get Zero() As Number
    Set Zero = Wrap(0)
End Property
Public Property Get One() As Number
    Set One = Wrap(1)
End Property
Public Property Get Two() As Number
    Set Two = Wrap(2)
End Property
Public Property Get MinusOne() As Number
    Set MinusOne = Wrap(-1)
End Property
Public Property Get Infinite() As Number
    Set Infinite = Wrap(-1)
End Property

''<default/>
Public Function Create(ByRef Expression As Variant) As Number
Attribute Create.VB_UserMemId = 0
    If Info.IsNumber(Expression) Then
        Set Create = Wrap(Expression)
    Else
        
    End If
End Function
Public Function Wrap(ByVal StdNumberValue As Variant) As NumberStd
    Set Wrap = New NumberStd
    Call Wrap.Constructor0(StdNumberValue)
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

Public Function Add(ByVal Num1 As Number, ByVal Num2 As Number) As Number
    
End Function
Public Function Subtract(ByVal Num1 As Number, ByVal Num2 As Number) As Number
    
End Function
Public Function Multiply(ByVal Num1 As Number, ByVal Num2 As Number) As Number
    
End Function
Public Function Divide(ByVal Num1 As Number, ByVal Num2 As Number) As Number
    
End Function
Public Function Power(ByVal Num1 As Number, ByVal Num2 As Number) As Number
    
End Function

Public Function BitwiseNot(ByVal Number As Number) As Number
    
End Function
Public Function BitwiseAnd(ByVal Number As Number) As Number
    
End Function
Public Function BitwiseOr(ByVal Number As Number) As Number
    
End Function
Public Function BitwiseXor(ByVal Number As Number) As Number
    
End Function
Public Function Negate(ByVal Number As Number) As Number
    
End Function

'*********************************************
'*********************************************
'
'</section>
'---------------------------------------------
'---------------------------------------------
