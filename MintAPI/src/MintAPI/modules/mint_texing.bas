Attribute VB_Name = "mint_texing"
Option Explicit

'by Ali Mousavi Kherad (LGPL-v3)
'Free to use and distribute but including my name and email as (alimousavikherad@gmail.com)!

'MintAPI Texing is a technology that i personaly featured (Ali Mousavi Kherad),
'It most like SOAP on net but designed for both local and network applications.
'Texing provides some powerfull activities such as saving a whole class to a
'File or fill it to a stream class or event put it in processes.
'On the network this work like SOAP but also you can control it by your self as
'You wish.
'

'Here is a simple output of String
'IClassTexerS0.1<Action,com.mintapi.Action/1.0.0>[[classparameters]]{ //static class parameters
'    IsAutomated:"true";
'    Name:"rename_filesAction";
'    Script:ARRAY(Variant)[[2 - first bound[,2 nd Bound[,...]]]]{ //mitone chand bodi bashe!
'        "C:\Ali.txt|C:\Renamed.txt",
'        OBJECT_STR([345345{< it is a address} or Language(""){< it is a class name or alias provides by HostProvider class}]);
'        OBJECT_STR(ClassProvider::com.myDll.myClass/0.1)script{
'            Property01:"it loads a client application implemented class.";
'            //ClassProvider::com.myDll.myClass/0.1
'            //It will ask MintAPI.ClassProvider engine to load a
'            //myClass version 0.1 from identified application or
'            //Dll com.myDll as a class alias string.
'            //This alias contains:
'            //ProviderEngine::[[projectpackage.]projectname.]Class[/classversion]
'        }
'    };
'    obj:OBJECT_STR();
'  //ex:
'}

Public Const CLASSTEXER_NAME As String = "MClassTexer"
Public Const CLASSTEXER_MAJVERSION As Long = 0
Public Const CLASSTEXER_MINVERSION As Long = 1
Public Const CLASSTEXER_VERSION As Long = (CLASSTEXER_MAJVERSION * &H100) + CLASSTEXER_MINVERSION '0x 000100 {00.01.00 ~ 99.99.00}

Public Const INFINITE_CLASS_PARAMETERS As Long = -1

Public Enum mint_TexingFormat
    tf_Stream = 0
    
    tf_Readable = 1 'can be readed by user.
    
    tf_MultiLine = 2 'is multiline.
    
    tf_NoIndented = 0
    tf_1Indented = 4 'this activates under tfMultiline state.
    tf_4Indented = &HC
    
    tf_Auto_Editable_NoIndent = tf_Readable Or tf_MultiLine
    tf_Auto_Editable_Compressed = tf_Auto_Editable_NoIndent Or tf_1Indented
    tf_Auto_Editable = tf_Auto_Editable_NoIndent Or tf_4Indented
    tf_Auto_Binary = tf_Stream
End Enum

Public Type TEX_ROW
    ID As String
    Value As Variant
End Type
Public Type TEX_CONTENT
    rowsCount As Long
    rows() As TEX_ROW

    ClassName As String
    CLASSALIAS As String
    CLASSVERSIONSTRING As String
    classParameters As Long
End Type


Public Function createClassData(ClassName As String, Optional classParameters As Long = INFINITE_CLASS_PARAMETERS _
                                ) As TEX_CONTENT

    createClassData.ClassName = ClassName
    createClassData.classParameters = classParameters
End Function
Public Sub appendClassData(classData As TEX_CONTENT, ID As String, Value)
    Dim cCount As Long
    cCount = classData.rowsCount
    ReDim Preserve classData.rows(cCount)
    classData.rows(cCount).ID = ID
    If varType(Value) = VBObject Then
        Set classData.rows(cCount).Value = Value
    Else
            classData.rows(cCount).Value = Value
    End If
End Sub

Public Function mint_tex_createheader(texType As mint_TexingType, ClassName As String, CLASSALIAS As String, CLASSVERSIONSTRING As String, ClassParametersCount As Long) As String
    Dim retP1 As String, retP2 As String
    retP1 = CLASSTEXER_NAME & IIf(texType = tt_Binary, "B", "S") & CLASSTEXER_MAJVERSION & "." & CLASSTEXER_MINVERSION
    retP2 = "<" & ClassName & "," & CLASSALIAS & "/" & CLASSVERSIONSTRING & ">"
    
    If ClassParametersCount > 0 Then
        mint_tex_createheader = retP1 & retP2 & "[" & ClassParametersCount & "]"
    Else
        mint_tex_createheader = retP1 & retP2
    End If
End Function


Public Function create_array_string() As String
    'arrayname : { data1 , data2 }
End Function
Public Function create_array_bytearray() As Byte()
    'arrayname : {binary1,binary2}
End Function

Private Function mint_tex_get_nameof_string(strTextData As String) As String
    
End Function
Private Function mint_tex_get_typeof_string(strTextData As String) As String
    
End Function

Private Function mint_tex_create_texer_header(classData As TEX_CONTENT) As String
    Dim str_tex_header As String, str_buffer As String, param_str As String
    'str_tex_header = "IClassTexer<"
    'IClassTexer<Name,Alias/Version>[...]
    str_buffer = classData.ClassName & "," & _
                 classData.CLASSALIAS & "/" & _
                 classData.CLASSVERSIONSTRING
    'a means class only have one parameter.
    'i means infinit parameters.
    'n no parameters.
    'else number of params.
    
    Select Case classData.classParameters
        Case -1 'infinit params
            param_str = "i"
        Case 0 'no param
            param_str = "n"
        Case 1 'one param
            param_str = "a"
        Case Else 'multi params
            param_str = CStr(classData.classParameters)
    End Select
    str_tex_header = _
    "IClassTexer<" & str_buffer & ">[" & param_str & "]"
    
    mint_tex_create_texer_header = str_tex_header
End Function
Public Function mint_totex_bytearray(classData As TEX_CONTENT, Optional texing_format As mint_TexingFormat) As Byte()
On Error GoTo Err
    Dim header As String, exContent As String, bufContent As String
    
    header = mint_tex_create_texer_header(classData)

    Dim i As Long, vt As VbVarType
    For i = 0 To classData.rowsCount - 1
        vt = varType(classData.rows(i).Value)

        If vt = vbArray Then

        ElseIf vt = VBObject Then
            'bufContent = classData.rows(i).ID & ":" & PrependTabToNewLines(fclass(classData.rows(i).Value).toString())
        ElseIf vt = vbUserDefinedType Then

        ElseIf vt = vbEmpty Or vt = vbNull Then

        Else
            GoTo Err
        End If

        exContent = exContent & (vbCrLf & vbTab & bufContent)
    Next

    mint_totex_bytearray = exContent & vbCrLf & "}"
    Exit Function
Err:
    throw Exception("at<mintapi.pubs.mint_totex_bytearray()> : Some classData As TEX_RECORDS Rows Values Type Are Unknown.")
End Function
Public Function mint_totex_string(classData As TEX_CONTENT, Optional texing_format As mint_TexingFormat) As String
On Error GoTo Err
    Dim header As String, exContent As String, bufContent As String
    
    header = mint_tex_create_texer_header(classData)

    Dim i As Long, vt As VbVarType
    For i = 0 To classData.rowsCount - 1
        vt = varType(classData.rows(i).Value)

        If vt = vbArray Then

        ElseIf vt = VBObject Then
            'bufContent = classData.rows(i).ID & ":" & PrependTabToNewLines(fclass(classData.rows(i).Value).toString())
        ElseIf vt = vbUserDefinedType Then

        ElseIf vt = vbEmpty Or vt = vbNull Then

        Else
            GoTo Err
        End If

        exContent = exContent & (vbCrLf & vbTab & bufContent)
    Next

    mint_totex_string = exContent & vbCrLf & "}"
    Exit Function
Err:
    throw Exception("at<mintapi.pubs.mint_totex_string()> : Some classData As TEX_RECORDS Rows Values Type Are Unknown.")
End Function

Public Function AppendTabToNewLines(str As String) As String
    AppendTabToNewLines = vbTab & Replace(str, vbCrLf, vbCrLf & vbTab)
End Function

Public Function GetObjectXML(texedObject As texedObject) As String
    If texedObject Is Nothing Then throw ArgumentNullException
    If Not TypeOf texedObject Is ITexable Then throw InvalidArgumentTypeException
    Dim texS As String
    Dim xml_header_retVal As String, xml_Content As String
    
    Dim objName As String, objType As String
    Dim objVersion As Long
    
    objName = texedObject.Name
    objType = texedObject.TypeName
    
    xml_header_retVal = "<object name=""" & objName & """ vtype=""class"" type=""" & objType & """ version=""" & objVersion & """>"
    
    xml_Content = " content "
    
    GetObjectXML = xml_header_retVal & xml_Content & "</object>"
End Function
