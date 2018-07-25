VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DateTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'@PROJECT_LICENSE

Option Base 0
Option Explicit
Const CLASSID As String = "DateTime"
Const CLASSALIAS As String = APPLICATIONDOMAIN & "." & CLASSID
Const CLASSVERSIONSTRING As String = "0.1"
Const CLASSALIASSTRING As String = CLASSALIAS & "/" & CLASSVERSIONSTRING


Implements ICloneable
'Implements IAliasable

Dim d As Date

Friend Sub Initialize(Optional Value)
    If IsMissing(Value) Then
        d = VBA.Now
    Else
        If (VarType(Value) And vbDate) = vbDate Then
            d = Value
        Else 'is string
            d = CDate(Value)
        End If
    End If
End Sub
Friend Sub FromByteArrayData(B() As Byte)

End Sub

Public Function toSTDDate() As Date
    toSTDDate = d
End Function

Public Function Parse(parseContent) As DateTime
    Dim d As Date
    d = CDate(parseContent)
    Set Parse = New DateTime
    Call Parse.Initialize(d)
End Function
Public Function TryParse(parseContent, outValue As DateTime) As Boolean
On Error GoTo Err
    Dim d As Date
    d = CDate(parseContent)
    Set outValue = New DateTime
    Call outValue.Initialize(d)
    TryParse = True
    Exit Function
Err:
    TryParse = False
    Set outValue = Nothing
End Function

Public Function toLongDateString() As String
    toLongDateString = FormatDateTime(d, vbLongDate)
End Function
Public Function toLongTimeString() As String
    toLongTimeString = FormatDateTime(d, vbLongTime)
End Function
Public Function toLongDateTimeString() As String
    toLongDateTimeString = toLongDateString & " " & toLongTimeString
End Function
Public Function toShortDateString() As String
    toShortDateString = FormatDateTime(d, vbShortDate)
End Function
Public Function toShortTimeString() As String
    toShortTimeString = FormatDateTime(d, vbShortTime)
End Function
Public Function toShortDateTimeString() As String
    toShortDateTimeString = toShortDateString & " " & toShortTimeString
End Function

Public Function Format(strFormat As String) As String
    Format = VBA.Format$(CStr(d), strFormat)
End Function
Public Function ToString() As String
    ToString = CStr(d)
End Function


Public Property Get Calendar() As VbCalendar
    Calendar = VBA.Calendar
End Property
Public Property Let Calendar(Value As VbCalendar)
    VBA.Calendar = Value
End Property

Public Property Get IsLeapYear() As Boolean
    
End Property
Public Property Get Now() As DateTime
    Set Now = New DateTime
    Call Now.Initialize
End Property

Public Property Get Yesterday() As DateTime
    Set Yesterday = New DateTime
    Call Yesterday.Initialize(VBA.Now - 1)
End Property
Public Property Get Tomorrow() As DateTime
    Set Tomorrow = New DateTime
    Call Tomorrow.Initialize(VBA.Now + 1)
End Property

Public Property Get Year() As Long
    Year = VBA.Year(d)
End Property
Public Property Let Year(Value As Long)
    
End Property
Public Property Get Month() As Long
    Month = VBA.Month(d)
End Property
Public Property Let Month(Value As Long)
    
End Property
Public Property Get Day() As Long
    Day = VBA.Day(d)
End Property
Public Property Let Day(Value As Long)
    
End Property

Public Property Get Hour() As Long
    Hour = VBA.Hour(d)
End Property
Public Property Let Hour(Value As Long)

End Property
Public Property Get Minute() As Long
    Minute = VBA.Minute(d)
End Property
Public Property Let Minute(Value As Long)

End Property
Public Property Get Second() As Long
    Second = VBA.Second(d)
End Property
Public Property Let Second(Value As Long)

End Property

Public Property Get Week() As Long
    Week = VBA.Weekday(d)
End Property

Public Property Get DayName(dayIndex As Long) As String

End Property
Public Property Get MonthName(monthIndex As Long) As String

End Property

Public Function Compare(DateTimeLower As DateTime, Optional DateTimeHigher As DateTime = Nothing) As TristateValues
    If DateTimeHigher Is Nothing Then Set DateTimeHigher = Me
    Dim d1 As Date, d2 As Date
    d2 = DateTimeHigher.toSTDDate
    d1 = DateTimeLower.toSTDDate
    If d2 > d1 Then
        Compare = tvMore
    ElseIf d2 < d1 Then
        Compare = tvLess
    Else
        Compare = tvEqual
    End If
End Function

Private Function ICloneable_Clone() As Object
    Dim dt As New DateTime
    Call dt.Initialize(d)
    Set ICloneable_Clone = dt
End Function

Private Property Get IAliasable_Alias() As String
    IAliasable_Alias = CLASSALIASSTRING
End Property
