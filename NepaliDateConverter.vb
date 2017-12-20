Public Module NepaliDateConverter
    Public Enum DatePart
        YEAR
        MONTH
        DAY
        DAYNAME
        MONTHNAME
        WEEKDAY
    End Enum

    Private nepaliDateDictionary = New Dictionary(Of DatePart, String) From
        {{DatePart.YEAR, 0}, {DatePart.MONTH, 0}, {DatePart.DAY, 0}, {DatePart.DAYNAME, 0}, {DatePart.MONTHNAME, 0},
        {DatePart.WEEKDAY, 0}}
    Private englishDateDictionary = New Dictionary(Of DatePart, String) From
        {{DatePart.YEAR, 0}, {DatePart.MONTH, 0}, {DatePart.DAY, 0}, {DatePart.DAYNAME, 0}, {DatePart.MONTHNAME, 0},
        {DatePart.WEEKDAY, 0}}

    Private englishMonthLength = {31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31}
    Private englishLeapYearMonthLength = {31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31}
    Private ReadOnly nepaliMonths As String() = {"Baisakh", "Jestha", "Ashad", "Shrawn", "Bhadra",
        "Ashwin", "Kartik", "Mangshir", "Poush", "Magh", "Falgun", "Chaitra"}
    Private ReadOnly englishMonths As String() = {"January", "February", "March", "April", "May",
    "June", "July", "August", "September", "October", "November", "December"}
    Private ReadOnly daysOfWeek = {"Sunday", "Monday", "Tuesday", "Wednesday", "Thursday",
        "Friday"}
    Private ReadOnly bs()() As Integer = ({({2000, 30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31}),
            ({2001, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30}),
            ({2002, 31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30}),
            ({2003, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31}),
            ({2004, 30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31}),
            ({2005, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30}),
            ({2006, 31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30}),
            ({2007, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31}),
            ({2008, 31, 31, 31, 32, 31, 31, 29, 30, 30, 29, 29, 31}),
            ({2009, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30}),
            ({2010, 31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30}),
            ({2011, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31}),
            ({2012, 31, 31, 31, 32, 31, 31, 29, 30, 30, 29, 30, 30}),
            ({2013, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30}),
            ({2014, 31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30}),
            ({2015, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31}),
            ({2016, 31, 31, 31, 32, 31, 31, 29, 30, 30, 29, 30, 30}),
            ({2017, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30}),
            ({2018, 31, 32, 31, 32, 31, 30, 30, 29, 30, 29, 30, 30}),
            ({2019, 31, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31}),
            ({2020, 31, 31, 31, 32, 31, 31, 30, 29, 30, 29, 30, 30}),
            ({2021, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30}),
            ({2022, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 30}),
            ({2023, 31, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31}),
            ({2024, 31, 31, 31, 32, 31, 31, 30, 29, 30, 29, 30, 30}),
            ({2025, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30}),
            ({2026, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31}),
            ({2027, 30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31}),
            ({2028, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30}),
            ({2029, 31, 31, 32, 31, 32, 30, 30, 29, 30, 29, 30, 30}),
            ({2030, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31}),
            ({2031, 30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31}),
            ({2032, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30}),
            ({2033, 31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30}),
            ({2034, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31}),
            ({2035, 30, 32, 31, 32, 31, 31, 29, 30, 30, 29, 29, 31}),
            ({2036, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30}),
            ({2037, 31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30}),
            ({2038, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31}),
            ({2039, 31, 31, 31, 32, 31, 31, 29, 30, 30, 29, 30, 30}),
            ({2040, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30}),
            ({2041, 31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30}),
            ({2042, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31}),
            ({2043, 31, 31, 31, 32, 31, 31, 29, 30, 30, 29, 30, 30}),
            ({2044, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30}),
            ({2045, 31, 32, 31, 32, 31, 30, 30, 29, 30, 29, 30, 30}),
            ({2046, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31}),
            ({2047, 31, 31, 31, 32, 31, 31, 30, 29, 30, 29, 30, 30}),
            ({2048, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30}),
            ({2049, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 30}),
            ({2050, 31, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31}),
            ({2051, 31, 31, 31, 32, 31, 31, 30, 29, 30, 29, 30, 30}),
            ({2052, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30}),
            ({2053, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 30}),
            ({2054, 31, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31}),
            ({2055, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30}),
            ({2056, 31, 31, 32, 31, 32, 30, 30, 29, 30, 29, 30, 30}),
            ({2057, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31}),
            ({2058, 30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31}),
            ({2059, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30}),
            ({2060, 31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30}),
            ({2061, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31}),
            ({2062, 30, 32, 31, 32, 31, 31, 29, 30, 29, 30, 29, 31}),
            ({2063, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30}),
            ({2064, 31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30}),
            ({2065, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31}),
            ({2066, 31, 31, 31, 32, 31, 31, 29, 30, 30, 29, 29, 31}),
            ({2067, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30}),
            ({2068, 31, 31, 32, 32, 31, 30, 30, 29, 30, 29, 30, 30}),
            ({2069, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31}),
            ({2070, 31, 31, 31, 32, 31, 31, 29, 30, 30, 29, 30, 30}),
            ({2071, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30}),
            ({2072, 31, 32, 31, 32, 31, 30, 30, 29, 30, 29, 30, 30}),
            ({2073, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 31}),
            ({2074, 31, 31, 31, 32, 31, 31, 30, 29, 30, 29, 30, 30}),
            ({2075, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30}),
            ({2076, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 30}),
            ({2077, 31, 32, 31, 32, 31, 30, 30, 30, 29, 30, 29, 31}),
            ({2078, 31, 31, 31, 32, 31, 31, 30, 29, 30, 29, 30, 30}),
            ({2079, 31, 31, 32, 31, 31, 31, 30, 29, 30, 29, 30, 30}),
            ({2080, 31, 32, 31, 32, 31, 30, 30, 30, 29, 29, 30, 30}),
            ({2081, 31, 31, 32, 32, 31, 30, 30, 30, 29, 30, 30, 30}),
            ({2082, 30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 30, 30}),
            ({2083, 31, 31, 32, 31, 31, 30, 30, 30, 29, 30, 30, 30}),
            ({2084, 31, 31, 32, 31, 31, 30, 30, 30, 29, 30, 30, 30}),
            ({2085, 31, 32, 31, 32, 30, 31, 30, 30, 29, 30, 30, 30}),
            ({2086, 30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 30, 30}),
            ({2087, 31, 31, 32, 31, 31, 31, 30, 30, 29, 30, 30, 30}),
            ({2088, 30, 31, 32, 32, 30, 31, 30, 30, 29, 30, 30, 30}),
            ({2089, 30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 30, 30}),
            ({2090, 30, 32, 31, 32, 31, 30, 30, 30, 29, 30, 30, 30})})


    Public Function IsLeapYear(ByVal year As Integer) As Boolean
        If year Mod 100 = 0 Then
            If year Mod 400 = 0 Then Return True Else Return False
        Else
            If year Mod 4 = 0 Then Return True Else Return False
        End If
    End Function

    Public Function IsEnglishDateConvertable(ByVal year As Integer, ByVal month As Integer,
                                              ByVal day As Integer) As Boolean
        If (year < 1944 Or year > 2033) Or (month < 1 Or month > 12) Or
            (day < 1 Or day > If(IsLeapYear(year), englishLeapYearMonthLength(month - 1),
            englishMonthLength(month - 1))) Then Return False Else Return True
    End Function

    Public Function IsNepaliDateConvertable(ByVal year As Integer, ByVal month As Integer,
                                              ByVal day As Integer) As Boolean
        If (year < 2000 Or year > 2089) Or (month < 1 Or month > 12) Or
            (day < 1 Or day > bs(year - 2000)(month)) Then Return False Else Return True
    End Function

    Public Function ConvertEnglishToNepali(ByVal englishYear As Integer, ByVal englishMonth As Integer,
                                              ByVal englishDay As Integer) As Dictionary(Of DatePart, String)
        'setting the data by matching between two calendars
        Dim definedEnglishYear = 1944, totalEnglishDays = 0
        Dim definedNepaliYear = 2000, definedNepaliMonth = 9, definedNepaliDay = 17 - 1
        Dim codedMonthLength = 0
        'weekday represent the day in a week in terms of number
        Dim weekday = 7 - 1
        Dim nepaliYear = 0, nepaliMonth = 0, nepaliDay = 0

        'counting the total number of days starting from year 1944 as we have data for nepali calendar from that year
        'counting days from years falling between the given year and starting year
        For i = 0 To englishYear - definedEnglishYear - 1
            If IsLeapYear(definedEnglishYear + i) Then
                For j = 0 To 12 - 1
                    totalEnglishDays += englishLeapYearMonthLength(j)
                Next j
            Else
                For j = 0 To 12 - 1
                    totalEnglishDays += englishMonthLength(j)
                Next j
            End If
        Next i
        'counting days from the past month except the given month
        For i = 0 To (englishMonth - 1) - 1
            If IsLeapYear(englishYear) Then
                totalEnglishDays += englishLeapYearMonthLength(i)
            Else
                totalEnglishDays += englishMonthLength(i)
            End If
        Next i
        'adding remaining given days
        totalEnglishDays = totalEnglishDays + englishDay

        'moving a total number of calculated days forward from the matched day between calendars
        nepaliYear = definedNepaliYear
        nepaliMonth = definedNepaliMonth
        nepaliDay = definedNepaliDay
        For dayCounter = 1 To totalEnglishDays
            codedMonthLength = bs(nepaliYear - definedNepaliYear)(nepaliMonth)
            nepaliDay += 1
            weekday += 1
            If (nepaliDay > codedMonthLength) Then
                nepaliMonth += 1
                nepaliDay = 1
            End If
            If (weekday > 7) Then weekday = 1
            If (nepaliMonth > 12) Then
                nepaliYear += 1
                nepaliMonth = 1
            End If
        Next dayCounter
        nepaliDateDictionary(DatePart.YEAR) = nepaliYear
        nepaliDateDictionary(DatePart.MONTH) = nepaliMonth.ToString("D2")
        nepaliDateDictionary(DatePart.DAY) = nepaliDay.ToString("D2")
        nepaliDateDictionary(DatePart.DAYNAME) = daysOfWeek(weekday - 1)
        nepaliDateDictionary(DatePart.MONTHNAME) = nepaliMonths(nepaliMonth - 1)
        nepaliDateDictionary(DatePart.WEEKDAY) = weekday
        Return nepaliDateDictionary
    End Function

    Public Function ConvertNepaliToEnglish(ByVal nepaliYear As Integer, ByVal nepaliMonth As Integer,
                                      ByVal nepaliDay As Integer) As Dictionary(Of DatePart, String)
        'setting the data by matching between two calendars
        Dim definedEnglishYear = 1943, definedEnglishMonth = 4, definedEnglishDay = 14 - 1
        Dim definedNepaliYear = 2000, definedNepaliMonth = 1, definedNepaliDay = 1,
            totalNepaliDays = 0
        Dim codedMonthLength = 0
        Dim weekday = 4 - 1
        Dim englishYear = 0, englishMonth = 0, englishDay = 0
        Dim codedYearIndex = 0, codedMonthIndex = 0

        'counting the total number of days starting from year 1944 as we have data for nepali calendar from that year
        'counting days from years falling between the given year and starting year
        For codedYearIndex = 0 To nepaliYear - definedNepaliYear - 1
            For codedMonthIndex = 1 To 12
                totalNepaliDays += bs(codedYearIndex)(codedMonthIndex)
            Next codedMonthIndex
        Next codedYearIndex
        'counting days from the past month except the given month
        For codedMonthIndex = 1 To nepaliMonth - 1
            totalNepaliDays += bs(codedYearIndex)(codedMonthIndex)
        Next codedMonthIndex
        'adding remaining given days
        totalNepaliDays += nepaliDay

        'moving a total number of calculated days forward from the matched day between calendars
        englishYear = definedEnglishYear
        englishMonth = definedEnglishMonth
        englishDay = definedEnglishDay
        For dayCounter = 1 To totalNepaliDays
            If IsLeapYear(englishYear) Then
                codedMonthLength = englishLeapYearMonthLength(englishMonth - 1)
            Else
                codedMonthLength = englishMonthLength(englishMonth - 1)
            End If
            englishDay += 1
            weekday += 1
            If (englishDay > codedMonthLength) Then
                englishMonth += 1
                englishDay = 1
            End If
            If (weekday > 7) Then weekday = 1
            If (englishMonth > 12) Then
                englishYear += 1
                englishMonth = 1
            End If
        Next dayCounter

        englishDateDictionary(DatePart.YEAR) = englishYear
        englishDateDictionary(DatePart.MONTH) = englishMonth.ToString("D2")
        englishDateDictionary(DatePart.DAY) = englishDay.ToString("D2")
        englishDateDictionary(DatePart.DAYNAME) = daysOfWeek(weekday - 1)
        englishDateDictionary(DatePart.MONTHNAME) = englishMonths(englishMonth - 1)
        englishDateDictionary(DatePart.WEEKDAY) = weekday
        Return englishDateDictionary
    End Function

    'return english date in format dd/mm/yyyy
    Public Function ConvertNepaliToFormattedEnglishDate(ByVal nepaliYear As Integer, ByVal nepaliMonth As Integer,
                                      ByVal nepaliDay As Integer) As String
        Dim englishDateDictionary = NepaliDateConverter.ConvertNepaliToEnglish(nepaliYear, nepaliMonth, nepaliDay)
        Return (englishDateDictionary(DatePart.DAY) & "\" & englishDateDictionary(DatePart.MONTH) & "\" &
                    englishDateDictionary(DatePart.YEAR))
    End Function

    'return nepali date in format dd/mm/yyyy
    Public Function ConvertEnglishToFormattedNepaliDate(ByVal englishYear As Integer, ByVal englishMonth As Integer,
                                      ByVal englishDay As Integer) As String
        Dim nepaliDateDictionary = NepaliDateConverter.ConvertEnglishToNepali(englishYear, englishMonth, englishDay)
        Return (nepaliDateDictionary(DatePart.DAY) & "\" & nepaliDateDictionary(DatePart.MONTH) & "\" &
                    nepaliDateDictionary(DatePart.YEAR))
    End Function

End Module
