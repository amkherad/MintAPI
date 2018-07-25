Attribute VB_Name = "pubs"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Copyright (C) 2013 Ali Mousavi Kherad and/or other contributors.
'' Contact: alimousavikherad@gmail.com http://www.sourceforge.net/users/amgmail
''
'' This file is part of the MintAPI dll of the MintAPI Toolkit.
''
'' $_BEGIN_LICENSE:LGPL$
''
'' GNU Lesser General Public License Usage
'' Alternatively, this file may be used under the terms of the GNU Lesser
'' General Public License version 2.1 as published by the Free Software
'' Foundation and appearing in the file LICENSE.LGPL included in the
'' packaging of this file.  Please review the following information to
'' ensure the GNU Lesser General Public License version 2.1 requirements
'' will be met: http://www.gnu.org/licenses/old-licenses/lgpl-2.1.html.
''
'' In addition, as a special exception, MintAPI gives you certain additional
'' rights.  These rights are described in the MintAPI LGPL Exception
'' version 1.1, included in the file LGPL_EXCEPTION.txt in this package.
''
'' GNU General Public License Usage
'' Alternatively, this file may be used under the terms of the GNU
'' General Public License version 3.0 as published by the Free Software
'' Foundation and appearing in the file LICENSE.GPL included in the
'' packaging of this file.  Please review the following information to
'' ensure the GNU General Public License version 3.0 requirements will be
'' met: http://www.gnu.org/copyleft/gpl.html.
''
'' $_END_LICENSE$
''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'by Ali Mousavi Kherad (LGPL-v3)
'Free to use and distribute but including my name as Ali Mousavi Kherad and email as (alimousavikherad@gmail.com)!
'Public mint_api_dialogs_last_choose_file_read_only_flag_state As Boolean
'Public mint_api_console_is_breaked As Boolean

'Public mint_api_winsock_instances As Long
'Public mint_api_winsock_versionrequired_buffer As Long

''Structure used for manipulating linger option.
'Public Type API_LINGER
'    l_onoff As Integer ' option on/off
'    l_linger As Integer ' linger time
'End Type
'Public Type API_S_UN_B
'    s_b1 As Byte
'    s_b2 As Byte
'    s_b3 As Byte
'    s_b4 As Byte
'End Type
'Public Type API_S_UN_W
'    s_w1 As Integer
'    s_w2 As Integer
'End Type
'Public Type API_INTERNET_ADDR
''    S_un_b As API_S_UN_B
''    S_un_w  As API_S_UN_W
'    S_addr As API_S_UN_B
'End Type
'Public Type API_SOCKADDR_IN
'    sin_family As Integer
'    sin_port As Integer
'    sin_addr As API_INTERNET_ADDR
'    sin_zero As String * 8
'End Type
'Public Type API_SOCKADDR
'    sin_family As Integer
'    sin_port As Integer
'    sin_addr As Long
'    sin_zero As String * 8
'End Type
'
'Public Type API_HostEnt
'    h_name As String ': char far*
'    h_aliases() As String 'char far* far* : means: Array<String>
'    h_addrtype As Integer
'    h_length As Integer
'    h_addr_list() As String 'char far* far* : means: Array<String>
'End Type
'Private Const WSADESCRIPTION_LEN = 256
'Private Const WSASYS_STATUS_LEN = 128
'Public Type API_WSADATA
'    wVersion As Integer
'    wHighVersion As Integer
'    szDescription As String * WSADESCRIPTION_LEN
'    szSystemStatus As String * WSASYS_STATUS_LEN
'    iMaxSockets As Integer
'    iMaxUdpDg As Integer
'    lpVendorInfo As Long
'End Type
'
'Public Type API_IPv4Address
'    IP As Long '4Bytes
'End Type 'Total 4 Bytes
'Public Type API_IPv6Address
'    P1 As Long 'High 8Bytes
'    P2 As Long
'    P3 As Long
'    P4 As Long 'Low  8Bytes
'End Type 'Total 16 Bytes

'Public Function GlobalFilters(Optional Includes, Optional Excludes) As GlobalFilters
'    Dim RetVal As GlobalFilters
'    If IsMissing(Includes) Then
'        Dim p(1) As String
'        p(0) = "*": p(1) = "*.*"
'        Includes = p
'    End If
'    If VarType(Includes) = (vbArray Or vbString) Then
'        RetVal.IncludeTemplates = Includes
'    Else
'        ReDim RetVal.IncludeTemplates(0)
'        RetVal.IncludeTemplates(0) = CStr(Includes)
'    End If
'    If Not IsMissing(Excludes) Then
'        If VarType(Excludes) = (vbArray Or vbString) Then
'            RetVal.ExcludeTemplates = Excludes
'        Else
'            ReDim RetVal.ExcludeTemplates(0)
'            RetVal.ExcludeTemplates(0) = CStr(Excludes)
'        End If
'    End If
'    GlobalFilters = RetVal
'End Function
'Public Function FileFilters(Optional Includes, Optional Excludes) As GlobalFilters
'    Dim RetVal As GlobalFilters
'    If IsMissing(Includes) Then
'        Dim p(1) As String
'        p(0) = "*": p(1) = "*.*"
'        Includes = p
'    End If
'    If VarType(Includes) = (vbArray Or vbString) Then
'        RetVal.IncludeTemplates = Includes
'    Else
'        ReDim RetVal.IncludeTemplates(0)
'        RetVal.IncludeTemplates(0) = CStr(Includes)
'    End If
'    If Not IsMissing(Excludes) Then
'        If VarType(Excludes) = (vbArray Or vbString) Then
'            RetVal.ExcludeTemplates = Excludes
'        Else
'            ReDim RetVal.ExcludeTemplates(0)
'            RetVal.ExcludeTemplates(0) = CStr(Excludes)
'        End If
'    End If
'    FileFilters = RetVal
'End Function
'Public Function StringArray(ParamArray Str() As Variant) As String()
'    Dim i As Long, strSize As Long, RetVal() As String
'    On Error GoTo zeroLength
'    strSize = UBound(Str) - LBound(Str) + 1
'    If strSize > 0 Then
'        ReDim RetVal(strSize - 1)
'        For i = 0 To strSize - 1
'            RetVal(i) = CStr(Str(i))
'        Next
'    End If
'    StringArray = RetVal
'zeroLength:
'End Function
'Public Function IncludeFilter_single(fs() As String, Expression As String) As Boolean
'    Dim i As Long
'    For i = 0 To ArraySize(fs) - 1
'        If (Expression) Like CStr(fs(i)) Then
'            IncludeFilter_single = True
'            Exit Function
'        End If
'    Next
'    IncludeFilter_single = False
'End Function
'Public Function IsFilterIncluded(Filter As GlobalFilters, Expression As String) As Boolean
'    IsFilterIncluded = ((IncludeFilter_single(Filter.IncludeTemplates, Expression)) And (Not IncludeFilter_single(Filter.ExcludeTemplates, Expression)))
'End Function



'---------------------------------------------
'
'
'Public Function Array_String(ParamArray str() As Variant) As String()
'    Dim i As Long, strSize As Long, retVal() As String
'    On Error GoTo zeroLength
'    strSize = UBound(str) - LBound(str) + 1
'zeroLength:
'    If strSize > 0 Then
'        ReDim retVal(strSize - 1)
'        For i = 0 To strSize - 1
'            retVal(i) = CStr(str(i))
'        Next
'    End If
'    Array_String = retVal
'End Function
'Public Function Array_Object(ParamArray Objects() As Variant) As Object()
'    Dim i As Long, objSize As Long, retVal() As Object
'    On Error GoTo zeroLength
'    objSize = UBound(Objects) - LBound(Objects) + 1
'zeroLength:
'    If objSize > 0 Then
'        ReDim retVal(objSize - 1)
'        For i = 0 To objSize - 1
'            Set retVal(i) = Objects(i)
'        Next
'    End If
'    Array_Object = retVal
'End Function
'Public Function Array_Double(ParamArray Doubles() As Variant) As Double()
'    Dim i As Long, dblSize As Long, retVal() As Double
'    On Error GoTo zeroLength
'    dblSize = UBound(Doubles) - LBound(Doubles) + 1
'zeroLength:
'    If dblSize > 0 Then
'        ReDim retVal(dblSize - 1)
'        For i = 0 To dblSize - 1
'            retVal(i) = CDbl(Doubles(i))
'        Next
'    End If
'    Array_Double = retVal
'End Function
'Public Function Array_Single(ParamArray Singles() As Variant) As Single()
'    Dim i As Long, sngSize As Long, retVal() As Single
'    On Error GoTo zeroLength
'    sngSize = UBound(Singles) - LBound(Singles) + 1
'zeroLength:
'    If sngSize > 0 Then
'        ReDim retVal(sngSize - 1)
'        For i = 0 To sngSize - 1
'            retVal(i) = CSng(Singles(i))
'        Next
'    End If
'    Array_Single = retVal
'End Function
'Public Function Array_Long(ParamArray Longs() As Variant) As Long()
'    Dim i As Long, lngSize As Long, retVal() As Long
'    On Error GoTo zeroLength
'    lngSize = UBound(Longs) - LBound(Longs) + 1
'zeroLength:
'    If lngSize > 0 Then
'        ReDim retVal(lngSize - 1)
'        For i = 0 To lngSize - 1
'            retVal(i) = CLng(Longs(i))
'        Next
'    End If
'    Array_Long = retVal
'End Function
'Public Function Array_Integer(ParamArray Ints() As Variant) As Integer()
'    Dim i As Long, intSize As Long, retVal() As Integer
'    On Error GoTo zeroLength
'    intSize = UBound(Ints) - LBound(Ints) + 1
'zeroLength:
'    If intSize > 0 Then
'        ReDim retVal(intSize - 1)
'        For i = 0 To intSize - 1
'            retVal(i) = CLng(Ints(i))
'        Next
'    End If
'    Array_Integer = retVal
'End Function
'Public Function Array_Byte(ParamArray Bytes() As Variant) As Byte()
'    Dim i As Long, btSize As Long, retVal() As Byte
'    On Error GoTo zeroLength
'    btSize = UBound(Bytes) - LBound(Bytes) + 1
'zeroLength:
'    If btSize > 0 Then
'        ReDim retVal(btSize - 1)
'        For i = 0 To btSize - 1
'            retVal(i) = CByte(Bytes(i))
'        Next
'    End If
'    Array_Byte = retVal()
'End Function
