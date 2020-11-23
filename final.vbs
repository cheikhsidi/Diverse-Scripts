
Option Explicit
'ThisFunction claculate the first Monday for each week for a whole year
Function WeekStartDate(intWeek As Integer, ByVal intYear As Integer, Optional intMonth As Integer = 1, Optional intDay As Integer = 1)

'Declaring variables
Dim FromDate As Date, lngAdd As Long
Dim WKDay, WDays As Integer

WDays = 0

'Checking that year should not have negative value
If intYear < 1 Then
    WeekStartDate = "Year cann't have negative value"
    Exit Function
End If

'Calculating the date
FromDate = DateSerial(intYear, intMonth, intDay)

'Getting the week day of the specified date considering monday as first day
WKDay = Weekday(FromDate, vbMonday)

'If value of week day is less than 4 then subtracting 1 from the week number
If WKDay > 4 Then
    WDays = (7 * intWeek) - WKDay + 1
Else
    WDays = (7 * (intWeek - 1)) - WKDay + 1
End If

'Return the first day of the week
WeekStartDate = FromDate + WDays
    
End Function
'This Function claculate all dates between two dates
Function getDates(ByVal StartDate As Date, ByVal EndDate As Date) As Variant

    Dim varDates()      As Date
    Dim lngDateCounter  As Long

    ReDim varDates(1 To CLng(EndDate) - CLng(StartDate))

    For lngDateCounter = LBound(varDates) To UBound(varDates)
        varDates(lngDateCounter) = CDate(StartDate)
        StartDate = CDate(CDbl(StartDate) + 1)
    Next lngDateCounter

    getDates = varDates

ClearMemory:
    If IsArray(varDates) Then Erase varDates
    lngDateCounter = Empty

End Function
'Adding Hours
Function getHours(ByVal sh As Worksheet, ByVal col As Integer)
    Dim intIncr As Integer
    Dim intCellCnt, i As Integer
    Dim datDate As Date
    
    intIncr = 30                                    'minutes to add each cell
    intCellCnt = 27                   '24h * 60m = 1440 minutes per day
    datDate = CDate("01/11/2013 08:00:00 AM")          'start date+time for first cell

    For i = 1 To intCellCnt                         'loop through n cells
        sh.Cells(i + 3, col) = Format(datDate, "h:mm AM/PM")    'write and format result
         sh.Cells(i + 3, col).NumberFormat = "h:mm AM/PM;@"
         datDate = DateAdd("n", intIncr, datDate)    'add increment value
    Next i
End Function
Function A_SelectAllMakeTable2()
    Dim tbl As ListObject
    
    Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$R$31"), , xlYes)
    tbl.TableStyle = "TableStyleMedium16"
End Function
Sub YearWorkbook1()
    Dim iWeek As Integer
    Dim sht As Variant
    Dim Sdt As Date
    Dim Edt As Date
    Dim num As Long
    Dim yr, i As Long
    Dim myArray As Variant
    Dim tbl As ListObject
    Dim answer As Integer
    
    
    yr = InputBox("Please Enter the year (YYYY)")
    Application.ScreenUpdating = False
   If Application.Sheets.Count > 1 Then
            answer = MsgBox("Oops You have more than one Sheet in this Workbookdo you want to continue?", vbYesNo + vbQuestion, "Generate Calendar for The Year " & yr)
    If answer = vbYes Then
    
   
    'Worksheets.Add After:=Worksheets(Worksheets.Count), _
      Count:=(52 - Worksheets.Count)
    iWeek = 1
    num = Day(WeekStartDate(iWeek + 1, yr)) - 1
    For Each sht In Worksheets
         Worksheets(sht.Name).Activate
        Sdt = WeekStartDate(iWeek, yr)
        Edt = DateAdd("ww", iWeek - 1, DateSerial(yr, 1, num))
        ' getting all dates in a week
       myArray = getDates(Sdt, Edt)
       'renaming each sheet
        sht.Name = Left(MonthName(Month(Sdt)), 3) + " " & Format(Day(Sdt), "00") + "-" & Left(MonthName(Month(Edt)), 3) + " " & Format(Day(Edt), "00")
   
        'Styling the Worksheet
       If sht.ListObjects.Count > 0 Then
        sht.ListObjects(1).Unlist
        sht.UsedRange.ClearFormats
       End If
       ActiveWorkbook.Sheets(sht.Name).ListObjects.Add(xlSrcRange, ActiveWorkbook.Sheets(sht.Name).Range("$A$1:$R$31"), , xlYes).TableStyle = "TableStyleMedium16"
        sht.Range("A1") = "APPT"
        sht.Range("B1") = "MONDAY"
        sht.Range("C1") = myArray(1)
        sht.Range("D1") = "APPT"
        sht.Range("E1") = "TUESDAY"
        sht.Range("F1") = myArray(2)
        sht.Range("G1") = "APPT"
        sht.Range("H1") = "WEDNESDAY"
        sht.Range("I1") = myArray(3)
        sht.Range("J1") = "APPT"
        sht.Range("K1") = "THURSDAY"
        sht.Range("L1") = myArray(4)
        sht.Range("M1") = "APPT"
        sht.Range("N1") = "FRIDAY"
        sht.Range("O1") = myArray(5)
        sht.Range("P1") = "APPT"
        sht.Range("Q1") = "SATURDAY"
        sht.Range("R1") = myArray(6)
           For i = 1 To 18 Step 3
                sht.Cells(2, i) = "TECHS:"
                sht.Cells(2, i).Font.Bold = True
                sht.Cells(3, i) = "TIME"
                sht.Cells(3, i).Font.Bold = True
                Call getHours(sht, i)
            Next i
        iWeek = iWeek + 1
    Next sht
    Application.ScreenUpdating = True
    Else
        'DoNothing
    End If
    End If
    
    
End Sub






