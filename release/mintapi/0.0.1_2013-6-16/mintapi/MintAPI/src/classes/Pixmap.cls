VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Pixmap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Base 0
Option Explicit
Const CLASSID As String = "Pixmap"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


Implements IAliasable


Public Enum ImageLoadFlags
    ilfTransparency = &H78000000
    ilfCopyNew = &H70000000 'default
    ilfBounded = &H71000000
    ilfChangeTo1bit = &H70000000 '2^x bits : x = 0
    ilfChangeTo2bit = &H70100000 '2^x bits : x = 1
    ilfChangeTo4bit = &H70200000 '2^x bits : x = 2
    ilfChangeTo8bit = &H70300000 '2^x bits : x = 3
    ilfChangeTo16bit = &H70400000 '2^x bits : x = 4
    ilfChangeTo32bit = &H70500000 '2^x bits : x = 5
    ilfChangeTo64bit = &H70600000 '2^x bits : x = 6
End Enum
Public Enum ImageResizeMethod
    irmDefault
End Enum
Private Const DEFAULTIMAGELOADFLAGS As Long = ilfTransparency Or ilfCopyNew

Public Enum ImageCompressionMethods
    icmJPG
    icmJPEG = icmJPG
    icmPNG
    icmBMP
    icmGIF
    icmTIF
    icmICO
    icmICON = icmICO
End Enum
Private Const DEFAULTIMAGECOMPRESSIONMETHOD As Long = icmJPG

Dim hdc As Long
Dim C() As Byte
Dim imgPath As String
Dim c_boundedVarPtr As Long

Private Sub Class_Initialize()
'
End Sub
Private Sub Class_Terminate()
    Call UnBoundArray
End Sub

Friend Sub Initialize(Optional Source)
    
End Sub

Public Function CreatePixmap(Width As Long, Height As Long, Optional BitsPerPixel As Long = 3, Optional ByVal ColorMode As ColorModes = cmRGB) As Pixmap
    
End Function

Public Function LoadFromFile(Path As String, Optional LoadMethod As ImageLoadFlags = DEFAULTIMAGELOADFLAGS) As Pixmap
    
End Function
Public Function LoadFromArray(targetArray, Optional LoadMethod As ImageLoadFlags = DEFAULTIMAGELOADFLAGS) As Pixmap
    
End Function
Public Function LoadFromhDC(hdc As Long, Optional LoadMethod As ImageLoadFlags = DEFAULTIMAGELOADFLAGS) As Pixmap
    
End Function
Public Function LoadFromByteArray(targetByteArray As ByteArray, Optional LoadMethod As ImageLoadFlags = DEFAULTIMAGELOADFLAGS) As Pixmap
    
End Function
Public Function LoadFromIPicture(pic As IPicture, Optional LoadMethod As ImageLoadFlags = DEFAULTIMAGELOADFLAGS) As Pixmap
    
End Function

Public Function Duplicate() As Pixmap
    Dim p As New Pixmap

End Function

Public Sub CopyToArray(targetArray)
    
End Sub
Public Sub BoundToArray(targetArray() As Byte)
    If c_boundedVarPtr <> 0 Then throw InvalidStatusException("Another Array Is Bounded To Pixmap.")
    Erase targetArray
    Call API_CopyMemory(ByVal API_VarPtrArray(targetArray), ByVal API_VarPtrArray(C), 4)
End Sub
Public Sub UnBoundArray()
    If c_boundedVarPtr = 0 Then Exit Sub
    Call API_ZeroMemory(ByVal c_boundedVarPtr, 4)
End Sub

Public Function toByteArray() As ByteArray
    
End Function
Public Sub fromByteArray(targetByteArray As ByteArray)
    
End Sub

Public Property Get Width() As Long
    
End Property
Public Property Get Height() As Long
    
End Property

Public Property Get BitsPerPixel() As Long
    
End Property
Public Property Get ColorMode() As ColorModes
    
End Property

Public Function GetPixel(X As Long, Y As Long) As Long
    
End Function
Public Sub SetPixel(X As Long, Y As Long, Value As Long)
    
End Sub

Public Sub Render(renderTargetObject As Object)
    Dim OB As ObjectBuffer
    Set OB = renderTargetObject
    Call OB.thrownot("rendertarget", 1)
End Sub

Public Function Resize(Width As Long, Height As Long, ResizeMethod As ImageResizeMethod) As Pixmap
    
End Function

Public Sub SaveToFile(Path As String, Optional Format As ImageCompressionMethods = DEFAULTIMAGECOMPRESSIONMETHOD)
    
End Sub

Public Function CreateGraphics() As Graphics
    
End Function


Public Property Get Gamma() As Long
    
End Property
Public Property Let Gamma(Value As Long)
    
End Property

Public Sub SetLevel(Value As Long)
    
End Sub

Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property
