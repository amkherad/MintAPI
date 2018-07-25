VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Graphics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Graphics is basic graphic object to paint on graphic devices."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Base 0
Option Explicit
Const CLASSID As String = "Graphics"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


Implements IAliasable


Public Enum Colors
    cActiveBorder = &H8000000A
    cActiveCaption = &H80000009
    cActiveCaptionText = &H0
    cAppWorkspace = &H8000000C
    cControl = &H8000000F
    cControlText = &H80000012
    cDesktop = &H80000001
    cGrayText = &H80000011
    cHighlight = &H8000000D
    cHighlightText = &H8000000E
    cInactiveBorder = &H8000000B
    cInactiveCaption = &H80000003
    cInactiveCaptionText = &H80000013
    cInfo = &H80000018
    cInfoText = &H80000017
    cMenu = &H80000004
    cMenuText = &H80000007
    cScrollBar = &H80000000
    cWindow = &HFFFFFF
    cWindowFrame = &H80000006
    cWindowText = &H80000008
    cTransparent = &HFFFFFFFF

    '----------------------------------------------

    cAliceBlue = &HFFF8F0
    cAntiqueWhite = &HD7EBFA
    cAqua = &HFFFF00
    cAquamarine = &HD4FF7F
    cAzure = &HFFFFF0
    cBeige = &HDCF5F5
    cBisque = &HC4E4FF
    cBlack = &H0
    cBlanchedAlmond = &HCDEBFF
    cBlue = &HFF0000
    cBlueViolet = &HE22B8A
    cBrown = &H2A2AA5
    cBurlyWood = &H87B8DE
    cCadetBlue = &HA09E5F
    cChartreuse = &HFF7F
    cChocolate = &H1E69D2
    cCoral = &H507FFF
    cCornflowerBlue = &HED9564
    cCornsilk = &HDCF8FF
    cCrimson = &H3C14DC
    cCyan = &HFFFF00
    cDarkBlue = &H8B0000
    cDarkCyan = &H8B8B00
    cDarkGoldenrod = &HB86B8
    cDarkGray = &HA9A9A9
    cDarkGreen = &H6400
    cDarkKhaki = &H6BB7BD
    cDarkMagenta = &H8B008B
    cDarkOliveGreen = &H2F6B55
    cDarkOrange = &H8CFF
    cDarkOrchid = &HCC3299
    cDarkRed = &H8B
    cDarkSalmon = &H7A96E9
    cDarkSeaGreen = &H8BBC8F
    cDarkSlateBlue = &H8B3D48
    cDarkSlateGray = &H4F4F2F
    cDarkTurquoise = &HD1CE00
    cDarkViolet = &HD30094
    cDeepPink = &H9314FF
    cDeepSkyBlue = &HFFBF00
    cDimGray = &H696969
    cDodgerBlue = &HFF901E
    cFirebrick = &H2222B2
    cFloralWhite = &HF0FAFF
    cForestGreen = &H228B22
    cFuchsia = &HFF00FF
    cGainsboro = &HDCDCDC
    cGhostWhite = &HFFF8F8
    cGold = &HD7FF
    cGoldenrod = &H20A5DA
    cGray = &H808080
    cGreen = &H8000
    cGreenYellow = &H2FFFAD
    cHoneydew = &HF0FFF0
    cHotPink = &HB469FF
    cIndianRed = &H5C5CCD
    cIndigo = &H82004B
    cIvory = &HF0FFFF
    cKhaki = &H8CE6F0
    cLavender = &HFAE6E6
    cLavenderBlush = &HF5F0FF
    cLawnGreen = &HFC7C
    cLemonChiffon = &HCDFAFF
    cLightBlue = &HE6D8AD
    cLightCoral = &H8080F0
    cLightCyan = &HFFFFE0
    cLightGoldenrodYellow = &HD2FAFA
    cLightGray = &HD3D3D3
    cLightGreen = &H90EE90
    cLightPink = &HC1B6FF
    cLightSalmon = &H7AA0FF
    cLightSeaGreen = &HAAB220
    cLightSkyBlue = &HFACE87
    cLightSlateGray = &H998877
    cLightSteelBlue = &HDEC4B0
    cLightYellow = &HE0FFFF
    cLime = &HFF00
    cLimeGreen = &H32CD32
    cLinen = &HE6F0FA
    cMagenta = &HFF00FF
    cMaroon = &H80
    cMediumAquamarine = &HAACD66
    cMediumBlue = &HCD0000
    cMediumOrchid = &HD355BA
    cMediumPurple = &HDB7093
    cMediumSeaGreen = &H71B33C
    cMediumSlateBlue = &HEE687B
    cMediumSpringGreen = &H9AFA00
    cMediumTurquoise = &HCCD148
    cMediumVioletRed = &H8515C7
    cMidnightBlue = &H701919
    cMintCream = &HFAFFF5
    cMistyRose = &HE1E4FF
    cMoccasin = &HB5E4FF
    cNavajoWhite = &HADDEFF
    cNavy = &H800000
    cOldLace = &HE6F5FD
    cOlive = &H8080
    cOliveDrab = &H238E6B
    cOrange = &HA5FF
    cOrangeRed = &H45FF
    cOrchid = &HD670DA
    cPaleGoldenrod = &HAAE8EE
    cPaleGreen = &H98FB98
    cPaleTurquoise = &HEEEEAF
    cPaleVioletRed = &H9370DB
    cPapayaWhip = &HD5EFFF
    cPeachPuff = &HB9DAFF
    cPeru = &H3F85CD
    cPink = &HCBC0FF
    cPlum = &HDDA0DD
    cPowderBlue = &HE6E0B0
    cPurple = &H800080
    cRed = &HFF
    cRosyBrown = &H8F8FBC
    cRoyalBlue = &HE16941
    cSaddleBrown = &H13458B
    cSalmon = &H7280FA
    cSandyBrown = &H60A4F4
    cSeaGreen = &H578B2E
    cSeaShell = &HEEF5FF
    cSienna = &H2D52A0
    cSilver = &HC0C0C0
    cSkyBlue = &HEBCE87
    cSlateBlue = &HCD5A6A
    cSlateGray = &H908070
    cSnow = &HFAFAFF
    cSpringGreen = &H7FFF00
    cSteelBlue = &HB48246
    cTan = &H8CB4D2
    cTeal = &H808000
    cThistle = &HD8BFD8
    cTomato = &H4763FF
    cTurquoise = &HD0E040
    cViolet = &HEE82EE
    cWheat = &HB3DEF5
    cWhite = &HFFFFFF
    cWhiteSmoke = &HF5F5F5
    cYellow = &HFFFF
    cYellowGreen = &H32CD9A
End Enum
Public Enum DrawHints
    dOpaqueRendering
    dAntialiasing
    dHighQualityAntialiasing
    dTextAntialiasing

    dNativeTextRenderer
    dGDIPlusTextRenderer

    dNativeRenderer
    dGDIPlusRenderer
End Enum
Public Enum ColorModes
    cmRGB
    cmARGB
    cmCMYK
    cmHSB
    cmLab
End Enum

Public Property Get Handle() As Long
    throw NotImplementedException
End Property
Public Property Get hdc() As Long
    throw NotImplementedException
End Property

Public Function CreateBackPixmap() As Pixmap
    throw NotImplementedException
End Function

Public Sub Clear(ByVal Color As Colors)
    throw NotImplementedException
End Sub
Public Sub Save()
    throw NotImplementedException
End Sub
Public Sub Restore()
    throw NotImplementedException
End Sub
Public Function toPixmap() As Pixmap
    throw NotImplementedException
End Function

Public Sub Invert()
    throw NotImplementedException
End Sub
Public Sub SetFilter(ByVal Filter As Object)
    throw NotImplementedException
End Sub

Public Property Get Width() As Long
    throw NotImplementedException
End Property
Public Property Get Height() As Long
    throw NotImplementedException
End Property
Public Property Get Dimenition() As Long
    throw NotImplementedException
End Property

Public Sub SetDrawHints(DrawHints As DrawHints, Optional Value As Boolean = True)
    throw NotImplementedException
End Sub

Public Sub DrawLines()
    throw NotImplementedException
End Sub
Public Sub DrawLine()
    throw NotImplementedException
End Sub

Public Sub DrawText()
    throw NotImplementedException
End Sub
Public Sub DrawSquare()
    throw NotImplementedException
End Sub
Public Sub DrawPath()
    throw NotImplementedException
End Sub
Public Sub DrawCircle()
    throw NotImplementedException
End Sub
Public Sub DrawOval()
    throw NotImplementedException
End Sub
Public Sub DrawTriangle()
    throw NotImplementedException
End Sub
Public Sub DrawCube()
    throw NotImplementedException
End Sub
Public Sub DrawGlobe()
    throw NotImplementedException
End Sub

Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property
