Attribute VB_Name = "windows_gls"
'
' Copyright (C) 2011 @t Team.
' All rights reserved.
' Contact: Ali Mousavi Kherad (alimousavikherad@gmail..com)
'
' This file is part of the baseMethods module of the @t Toolkit.
'
' GNU Lesser General Public License Usage
' This file may be used under the terms of the GNU Lesser General Public
' License version 2.1 as published by the Free Software Foundation and
' appearing in the file LICENSE.LGPL included in the packaging of this
' file. Please review the following information to ensure the GNU Lesser
' General Public License version 2.1 requirements will be met:
' http://www.gnu.org/licenses/old-licenses/lgpl-2.1.html.
'
' In addition, as a special exception, @t gives you certain additional
' rights. These rights are described in the @t LGPL Exception
' version 1.1, included in the file LGPL_EXCEPTION.txt in this package.
'
' GNU General Public License Usage
' Alternatively, this file may be used under the terms of the GNU General
' Public License version 3.0 as published by the Free Software Foundation
' and appearing in the file LICENSE.GPL included in the packaging of this
' file. Please review the following information to ensure the GNU General
' Public License version 3.0 requirements will be met:
' http://www.gnu.org/copyleft/gpl.html.
'
' Other Usage
' Alternatively, this file may be used in accordance with the terms and
' conditions contained in a signed written agreement between you and @t.
'
'-------------------------------------------------------------------------------------------

Option Explicit
Option Base 0
Const CLASSID As String = "windows_gls"

Public Type API_DwmMargins
    LeftWidth As Long
    RightWidth As Long
    TopHeight As Long
    BottomHeight As Long
End Type
Private Enum API_CompositionEnable
    Disable = 0
    Enable = 1
End Enum

Private Declare Function API_ExtendFrameIntoClientArea Lib "DwmApi" Alias "DwmExtendFrameIntoClientArea" (ByVal hwnd As Long, ByRef m As API_DwmMargins) As Long
Private Declare Function API_IsCompositionEnabled Lib "DwmApi" Alias "DwmIsCompositionEnabled" () As Boolean
Private Declare Function API_EnableComposition Lib "DwmApi" Alias "DwmEnableComposition" (compositionAction As API_CompositionEnable) As Long

Dim inited As Boolean

Public Sub Initialize()
    If inited Then Exit Sub
    Call STDCONSTS.Initialize
    Call Exceptions.Initialize
    inited = True
End Sub
Public Sub Dispose(Optional ByVal Force As Boolean = False)
    If Not inited Then Exit Sub
    Call STDCONSTS.Dispose(Force)
    inited = False
End Sub

Public Function IsVista7() As Boolean
    IsVista7 = GetVersion >= 5
End Function

Public Sub EnableVista7Composition()
    If Not IsVista7 Then throw OsNotSupported
    If Not API_IsCompositionEnabled Then Call API_EnableComposition(Enable)
End Sub
Public Sub DisableVista7Composition()
    If Not IsVista7 Then throw OsNotSupported
    If API_IsCompositionEnabled Then Call API_EnableComposition(Disable)
End Sub
Public Sub EnableWindowComposition(hwnd As Long, margins As API_DwmMargins)
    If Not IsVista7 Then throw OsNotSupported
    'If Not API_IsCompositionEnabled Then throw InvalidCallException(ERRORS_SYSTEMCOMPOSITONNOTENABLED)
    Call API_ExtendFrameIntoClientArea(hwnd, margins)
End Sub
Public Sub DisableWindowComposition(hwnd As Long)
    If Not IsVista7 Then throw OsNotSupported
    'If Not API_IsCompositionEnabled Then throw InvalidCallException(ERRORS_SYSTEMCOMPOSITONNOTENABLED)
    Dim m As API_DwmMargins
    Call API_ExtendFrameIntoClientArea(hwnd, m)
End Sub

Public Sub VBInitGalss(obj As Object, m As API_DwmMargins, ByVal BG As Long)
    If Not IsVista7 Then throw OsNotSupported
    'If Not API_IsCompositionEnabled Then throw InvalidCallException(ERRORS_SYSTEMCOMPOSITONNOTENABLED)
On Error GoTo err
    obj.BackColor = vbBlack
    Call VBExportMargin(obj, m, BG)
err:
End Sub
Public Sub VBFinalGlass(ByRef obj As Object)
    If Not IsVista7 Then throw OsNotSupported
    'If Not API_IsCompositionEnabled Then throw InvalidCallException(ERRORS_SYSTEMCOMPOSITONNOTENABLED)
    obj.BackColor = vbWhite
    Dim t As API_DwmMargins
    Call VBExportMargin(obj, t, vbWhite)
End Sub
Public Sub VBExcludeControlFromAeroGlass(ByRef obj As Object)
    If Not IsVista7 Then throw OsNotSupported
    'If Not API_IsCompositionEnabled Then throw InvalidCallException(ERRORS_SYSTEMCOMPOSITONNOTENABLED)
On Error GoTo err
If obj.Parent.BackColor = vbBlack Then
    Dim margins As API_DwmMargins
    margins.LeftWidth = obj.Left
    margins.RightWidth = IIf(obj.Parent.ScaleWidth - obj.Left - obj.Width > 0, obj.Parent.ScaleWidth - obj.Left - obj.Width, 0)
    margins.TopHeight = obj.Top
    margins.BottomHeight = IIf(obj.Parent.ScaleHeight - obj.Top - obj.Height > 0, obj.Parent.ScaleHeight - obj.Top - obj.Height, 0)
    obj.Parent.Line (obj.Left, obj.Top)-(obj.Left + obj.Width, obj.Top + obj.Height), vbWhite, BF
    Call API_ExtendFrameIntoClientArea(obj.Parent.hwnd, margins)
End If
err:
End Sub
Public Sub VBExportMargin(obj As Object, m As API_DwmMargins, ByVal bgcol As Long)
    If Not IsVista7 Then throw OsNotSupported
    'If Not API_IsCompositionEnabled Then throw InvalidCallException(ERRORS_SYSTEMCOMPOSITONNOTENABLED)
On Error GoTo err
    Call API_ExtendFrameIntoClientArea(obj.hwnd, m)
    Dim t As Integer
    t = obj.ScaleMode
    obj.ScaleMode = ScaleModeConstants.vbPixels
    If m.LeftWidth < obj.ScaleWidth - m.RightWidth And _
       m.TopHeight < obj.ScaleWidth - m.BottomHeight - 1 Then
        obj.Cls
        obj.Line (m.LeftWidth, m.TopHeight)-(obj.ScaleWidth - m.RightWidth - 1, obj.ScaleHeight - m.BottomHeight - 1), bgcol, BF
    End If
err:
On Error Resume Next
    obj.ScaleMode = t
End Sub

