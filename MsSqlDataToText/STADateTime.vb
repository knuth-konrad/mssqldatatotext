'------------------------------------------------------------------------------
'Purpose  : General STA specfifc (mostly tourism-related) helper methods
'
'Prereq.  : -
'Note     : -
'
'   Author: Knuth Konrad 09.04.2019
'   Source: -
'  Changed: -
'------------------------------------------------------------------------------
Imports System.ComponentModel

Imports libBAUtil.StringUtil

Namespace STA

   <DefaultProperty("Date")>
   <Serializable()>
   Public Class IATADateTime
      '------------------------------------------------------------------------------
      'Purpose  : Enhanced DateTime
      '
      'Prereq.  : -
      'Note     : Provides all of the standard DateTime properties, adds (touristic)
      '           properties/methods
      '
      '   Author: Knuth Konrad
      '     Date: 09.04.2019
      '   Source: -
      '  Changed: -
      '------------------------------------------------------------------------------

#Region "Declares"

      ' Format of IATA date, DDMMM or DDMMMYY
      Public Enum eIATADateType
         DateLong
         DateShort
      End Enum

      ' Casing of month abbreviations
      Public Enum eIATADateCasing
         ToLower  ' 01jan
         ToMixed  ' 01Jan
         ToUpper  ' 01JAN
      End Enum

      Private mdtmDate As DateTime

#End Region

#Region "Properties - Private"
#End Region

#Region "Properties - Public"

      ' Default property
      Public Property [Date] As DateTime
         Get
            Return mdtmDate
         End Get
         Set(value As DateTime)
            mdtmDate = value
         End Set
      End Property

      Public ReadOnly Property Day As Integer
         Get
            Return Me.Date.Day
         End Get
      End Property

      Public ReadOnly Property DayOfWeek As System.DayOfWeek
         Get
            Return Me.Date.DayOfWeek
         End Get
      End Property

      Public ReadOnly Property DayOfYear As Integer
         Get
            Return Me.Date.DayOfYear
         End Get
      End Property

      Public ReadOnly Property Hour As Integer
         Get
            Return Me.Date.Hour
         End Get
      End Property

      Public ReadOnly Property IATADateLong As String
         Get
            Return Me.ToDateIATA(eIATADateType.DateLong, eIATADateCasing.ToUpper)
         End Get
      End Property

      Public ReadOnly Property IATADateShort As String
         Get
            Return Me.ToDateIATA(eIATADateType.DateShort, eIATADateCasing.ToUpper)
         End Get
      End Property

      Public ReadOnly Property Millisecond As Integer
         Get
            Return Me.Date.Millisecond
         End Get
      End Property

      Public ReadOnly Property Minute As Integer
         Get
            Return Me.Date.Minute
         End Get
      End Property

      Public ReadOnly Property Month As Integer
         Get
            Return Me.Date.Month
         End Get
      End Property

      Public ReadOnly Property Now As DateTime
         Get
            Return DateTime.Now
         End Get
      End Property

      Public ReadOnly Property Second As Integer
         Get
            Return Me.Date.Second
         End Get
      End Property

      Public ReadOnly Property Ticks As Long
         Get
            Return Me.Date.Ticks
         End Get
      End Property

      Public ReadOnly Property TimeOfDay As System.TimeSpan
         Get
            Return Me.Date.TimeOfDay()
         End Get
      End Property

      Public ReadOnly Property Today As DateTime
         Get
            Return DateTime.Today
         End Get
      End Property

      Public ReadOnly Property UtcNow As DateTime
         Get
            Return DateTime.UtcNow
         End Get
      End Property

      Public ReadOnly Property Year As Integer
         Get
            Return Me.Date.Year
         End Get
      End Property

#End Region

#Region "Implementations"
#End Region

#Region "Methods - Private"
#End Region

#Region "Methods - Public"

      Public Function Add(ByVal timeSpan As System.TimeSpan) As DateTime
         Return Me.Date.Add(timeSpan)
      End Function

      Public Function AddDays(ByVal days As Double) As DateTime
         Return Me.Date.AddDays(days)
      End Function

      Public Function AddHours(ByVal hours As Double) As DateTime
         Return Me.Date.AddHours(hours)
      End Function

      Public Function AddMilliseconds(ByVal milliseconds As Double) As DateTime
         Return Me.Date.AddMilliseconds(milliseconds)
      End Function

      Public Function AddMinutes(ByVal minutes As Double) As DateTime
         Return Me.Date.AddMinutes(minutes)
      End Function

      Public Function AddMonths(ByVal months As Integer) As DateTime
         Return Me.Date.AddMonths(months)
      End Function

      Public Function AddSeconds(ByVal seconds As Double) As DateTime
         Return Me.Date.AddSeconds(seconds)
      End Function

      Public Function AddTicks(ByVal ticks As Long) As DateTime
         Return Me.Date.AddTicks(ticks)
      End Function

      Public Function AddYears(ByVal years As Integer) As DateTime
         Return Me.Date.AddYears(years)
      End Function

      Public Overloads Function CompareTo(ByVal value As DateTime) As Integer
         Return Me.Date.CompareTo(value)
      End Function

      Public Overloads Function CompareTo(ByVal value As Object) As Integer
         Return Me.Date.CompareTo(value)
      End Function

      Public Function DaysInMonth(ByVal year As Integer, ByVal month As Integer) As Integer
         Return DateTime.DaysInMonth(year, month)
      End Function

      Public Shadows Function Equals(ByVal value As DateTime) As Boolean
         Return Me.Date.Equals(value)
      End Function

      Public Function FromBinary(ByVal dateData As Long) As DateTime
         Return DateTime.FromBinary(dateData)
      End Function

      Public Function FromFileTime(ByVal dateData As Long) As DateTime
         Return DateTime.FromFileTime(dateData)
      End Function

      Public Function FromFileTimeUtc(ByVal dateData As Long) As DateTime
         Return DateTime.FromFileTimeUtc(dateData)
      End Function

      Public Function FromOADate(ByVal d As Double) As DateTime
         Return DateTime.FromOADate(d)
      End Function

      Public Overloads Function GetDateTimeFormats() As String()
         Return Me.Date.GetDateTimeFormats
      End Function

      Public Overloads Function GetDateTimeFormats(ByVal format As Char) As String()
         Return Me.Date.GetDateTimeFormats(format)
      End Function

      Public Overloads Function GetDateTimeFormats(ByVal format As Char, ByVal provider As System.IFormatProvider) As String()
         Return Me.Date.GetDateTimeFormats(format, provider)
      End Function

      Public Overloads Function GetDateTimeFormats(ByVal provider As System.IFormatProvider) As String()
         Return Me.Date.GetDateTimeFormats(provider)
      End Function

      Public Function IsDaylightSavingTime() As Boolean
         Return Me.Date.IsDaylightSavingTime
      End Function

      Public Function IsLeapYear(ByVal year As Integer) As Boolean
         Return DateTime.IsLeapYear(year)
      End Function

      Public Overloads Function Subtract(ByVal value As Date) As System.TimeSpan
         Return Me.Date.Subtract(value)
      End Function

      Public Overloads Function Subtract(ByVal value As System.TimeSpan) As DateTime
         Return Me.Date.Subtract(value)
      End Function

      Public Function ToDate(ByVal iataDate As String) As DateTime
         '------------------------------------------------------------------------------
         'Name     : ToDate
         'Purpose  : Converts a IATA date string to a DateTime type
         'Param    : iataDate - IATA date string, e.g. 01JAN or 01JAN19
         '
         'Prereq.  : -
         'Note     : -
         '
         '   Author: Knuth Konrad
         '     Date: 09.04.2019
         '   Source: -
         '  Changed: -
         '------------------------------------------------------------------------------

         ' Safe guards
         If iataDate.Length <> 5 AndAlso iataDate.Length <> 7 Then
            Throw New ArgumentOutOfRangeException("iataDate", "Invalid IATA date. Date must be either 5 or 7 characters, e.g. 01JAN or 01JAN19.")
         End If

         Dim sDay As String = String.Empty
         Dim sMonth As String = String.Empty, lMonth As Int32
         Dim sYear As String = String.Empty

         ' Short IATA date
         If iataDate.Length = 5 Then
            sDay = Left(iataDate, 2)
            sMonth = Right(iataDate, 3)
            sYear = DateTime.Now.Year.ToString
         ElseIf iataDate.Length = 7 Then
            sDay = Left(iataDate, 2)
            sMonth = iataDate.Substring(2, 3)
            sYear = Right(iataDate, 2)
         End If

         Select Case sMonth.ToLower
            Case "jan"
               lMonth = 1
            Case "feb"
               lMonth = 2
            Case "mar"
               lMonth = 3
            Case "apr"
               lMonth = 4
            Case "may"
               lMonth = 5
            Case "jun"
               lMonth = 6
            Case "jul"
               lMonth = 7
            Case "aug"
               lMonth = 8
            Case "sep"
               lMonth = 9
            Case "oct"
               lMonth = 10
            Case "nov"
               lMonth = 11
            Case "dec"
               lMonth = 12
            Case Else
               Throw New ArgumentOutOfRangeException("iataDate (month)", "Month must be notated as 3-letter English month abbreviations, e.g. OCT for October")
         End Select

         ' Should handle all other exceptions
         Return New DateTime(Integer.Parse(sYear), lMonth, Integer.Parse(sDay))

      End Function

      Public Function ToBinary() As Long
         Return Me.Date.ToBinary
      End Function

      Public Overloads Function ToDateIATA() As String

         ' Upper case is default
         Return Me.Date.ToString("ddMMM", Globalization.CultureInfo.InvariantCulture).ToUpper

      End Function

      Public Overloads Function ToDateIATA(ByVal dateType As eIATADateType) As String

         ' Upper case is default
         If dateType = eIATADateType.DateLong Then
            Return Me.Date.ToString("ddMMMyy", Globalization.CultureInfo.InvariantCulture).ToUpper
         Else
            Return Me.Date.ToString("ddMMM", Globalization.CultureInfo.InvariantCulture).ToUpper
         End If

      End Function

      Public Overloads Function ToDateIATA(ByVal nameCasing As eIATADateCasing) As String

         Select Case nameCasing
            Case eIATADateCasing.ToLower
               Return Me.Date.ToString("ddMMM", Globalization.CultureInfo.InvariantCulture).ToLower
            Case eIATADateCasing.ToMixed
               Return Me.Date.ToString("ddMMM", Globalization.CultureInfo.InvariantCulture)
            Case eIATADateCasing.ToUpper
               Return Me.Date.ToString("ddMMM", Globalization.CultureInfo.InvariantCulture).ToUpper
            Case Else
               Return Me.Date.ToString("ddMMM", Globalization.CultureInfo.InvariantCulture).ToUpper
         End Select

      End Function

      Public Overloads Function ToDateIATA(ByVal dateType As eIATADateType, ByVal nameCasing As eIATADateCasing) As String

         Dim sResult As String = Me.ToDateIATA(dateType)

         Select Case nameCasing
            Case eIATADateCasing.ToLower
               Return sResult.ToLower()
            Case eIATADateCasing.ToMixed
               Return sResult
            Case eIATADateCasing.ToUpper
               Return sResult.ToUpper
            Case Else
               Return sResult.ToUpper
         End Select

      End Function

      Public Function ToFileTime() As Long
         Return Me.Date.ToFileTime
      End Function

      Public Function ToFileTimeUtc() As Long
         Return Me.Date.ToFileTimeUtc
      End Function

      Public Function ToLocalTime() As Date
         Return Me.Date.ToLocalTime
      End Function

      Public Function ToLongDateString() As String
         Return Me.Date.ToLongDateString
      End Function

      Public Function ToLongTimeString() As String
         Return Me.Date.ToLongTimeString
      End Function

      Public Function ToOADate() As Double
         Return Me.Date.ToOADate
      End Function

      Public Function ToShortDateString() As String
         Return Me.Date.ToShortDateString
      End Function

      Public Function ToShortTimeString() As String
         Return Me.Date.ToShortTimeString
      End Function

      Public Overrides Function ToString() As String
         Return Me.Date.ToString
      End Function

      Public Overloads Function ToString(ByVal format As String) As String
         Return Me.Date.ToString(format)
      End Function

      Public Overloads Function ToString(ByVal format As String, ByVal provider As System.IFormatProvider) As String
         Return Me.Date.ToString(format, provider)
      End Function


#End Region

#Region "Constructor/Deconstructor"

      Public Sub New()
         Me.Date = New DateTime
      End Sub

      Public Sub New(ByVal newDate As DateTime)
         Me.Date = newDate
      End Sub

      Public Sub New(ByVal ticks As Int64)
         Me.Date = New DateTime(ticks)
      End Sub

      Public Sub New(ByVal ticks As Int64, ByVal kind As DateTimeKind)
         Me.Date = New DateTime(ticks, kind)
      End Sub

      Public Sub New(ByVal year As Int32, ByVal month As Int32, ByVal day As Int32)
         Me.Date = New Date(year, month, day)
      End Sub

      Public Sub New(ByVal year As Int32, ByVal month As Int32, ByVal day As Int32, ByVal calendar As Globalization.Calendar)
         Me.Date = New Date(year, month, day, calendar)
      End Sub

      Public Sub New(ByVal year As Int32, ByVal month As Int32, ByVal day As Int32,
                        ByVal hour As Int32, ByVal minute As Int32, second As Int32)
         Me.Date = New Date(year, month, day, hour, minute, second)
      End Sub

      Public Sub New(ByVal year As Int32, ByVal month As Int32, ByVal day As Int32,
                        ByVal hour As Int32, ByVal minute As Int32, second As Int32, kind As DateTimeKind)
         Me.Date = New Date(year, month, day, hour, minute, second, kind)
      End Sub

      Public Sub New(ByVal year As Int32, ByVal month As Int32, ByVal day As Int32,
                        ByVal hour As Int32, ByVal minute As Int32, second As Int32, ByVal calendar As Globalization.Calendar)
         Me.Date = New Date(year, month, day, hour, minute, second, calendar)
      End Sub

      Public Sub New(ByVal year As Int32, ByVal month As Int32, ByVal day As Int32,
                        ByVal hour As Int32, ByVal minute As Int32, second As Int32, ByVal millisecond As Int32)
         Me.Date = New Date(year, month, day, hour, minute, second, millisecond)
      End Sub

      Public Sub New(ByVal year As Int32, ByVal month As Int32, ByVal day As Int32,
                        ByVal hour As Int32, ByVal minute As Int32, second As Int32, ByVal millisecond As Int32,
                        kind As DateTimeKind)
         Me.Date = New Date(year, month, day, hour, minute, second, millisecond, kind)
      End Sub

      Public Sub New(ByVal year As Int32, ByVal month As Int32, ByVal day As Int32,
                        ByVal hour As Int32, ByVal minute As Int32, second As Int32, ByVal millisecond As Int32,
                        calendar As Globalization.Calendar)
         Me.Date = New Date(year, month, day, hour, minute, second, millisecond, calendar)
      End Sub

      Public Sub New(ByVal year As Int32, ByVal month As Int32, ByVal day As Int32,
                        ByVal hour As Int32, ByVal minute As Int32, second As Int32, ByVal millisecond As Int32,
                        calendar As Globalization.Calendar, kind As DateTimeKind)
         Me.Date = New Date(year, month, day, hour, minute, second, millisecond, calendar, kind)
      End Sub

      Public Sub New(ByVal iataDate As String)
         Me.Date = Me.ToDate(iataDate)
      End Sub

#End Region

   End Class

End Namespace
