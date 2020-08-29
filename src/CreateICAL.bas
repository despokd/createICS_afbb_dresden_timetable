'' Create a iCalendar from a timetable for AFBB trainees
''
'' @author Kilian Domaratius
'' @version 1.0

'global
Dim startDate As String
Dim endDate As String
Dim lastUpdate As String
Dim schoolClassName As String


Sub StartMacro_CreateICAL()
    'PURPOSE:   Determine how many seconds it took for code to completely run
    'SOURCE:    www.TheSpreadsheetGuru.com/the-code-vault
    
    Dim StartTime As Double
    Dim SecondsElapsed As Double
    
    'Remember time when macro starts
    StartTime = Timer
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   Start own code                                                       '                                                              '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    createICAL
    
    'Determine how many seconds code took to run
    SecondsElapsed = Round(Timer - StartTime, 2)
    
    'Notify user in seconds + additional info
    Dim Msg As String
        Msg = "iCal erfolgreich in " & SecondsElapsed & " Sekunden erstellt für " & schoolClassName & " - Stand: " & lastUpdate
    Dim ButtonStyle
        ButtonStyle = vbOKOnly
    Dim Title As String
        Title = schoolClassName & "-iCal"
    InformUser = MsgBox(Msg, ButtonStyle, Title)
    
End Sub

Sub setGlobals()

    startDate = getStartDate
    endDate = getEndDate
    lastUpdate = getLastUpdate
    schoolClassName = getSchoolClassName
    
End Sub

Sub createICAL()

'init
Dim ws As Worksheet
Set ws = Worksheets(1)
ws.Activate
setGlobals

Dim iCalString As String

'Create Calendar
iCalString = iCalString & "BEGIN:VCALENDAR" & insertNewLine
    iCalString = iCalString & "VERSION:2.0" & insertNewLine
    iCalString = iCalString & "PROID:-//tequilian//tequilian.de//DE" & insertNewLine 'change for author
    iCalString = iCalString & "METHOD:REQUEST" & insertNewLine 'alternative: PUBLISH
    iCalString = iCalString & "BEGIN:VTIMEZONE" & insertNewLine
    iCalString = iCalString & "TZID:Europe/Berlin" & insertNewLine
    iCalString = iCalString & "END:VTIMEZONE" & insertNewLine
    
    'Insert Events
    
    Dim rowNumber As Integer
    rowNumber = 6
    'Start by cell A6 as start date
    
    Dim dayNumber As Integer
    dayNumber = 0
    'Start day = startDate + 0 days
    
    Dim endFound As Boolean
    endFound = False
    
    
    Do While endFound = False
        If Range("A" & CStr(rowNumber)).Interior.Color <> "10079487" Then
            endFound = True
            
        ElseIf Range("A" & CStr(rowNumber)).Value <> "" Then
            'Day found
            iCalString = iCalString & createDay(rowNumber, dayNumber)
            dayNumber = dayNumber + 1
    
        End If
        
        'next row
        rowNumber = rowNumber + 1
    Loop
    
    'End insert Events
    
'End Calendar
iCalString = iCalString & "END:VCALENDAR" & insertNewLine

'let user save as iCal File
Dim saveDialog As Variant
saveDialog = Application.GetSaveAsFilename(InitialFileName:="Calendar-" & schoolClassName & "-" & (Format(CDate(lastUpdate), "ddmmyyyy")), FileFilter:="iCalendar(*.ics), *.ics")
Set fs = CreateObject("Scripting.FileSystemObject")
Set a = fs.CreateTextFile(saveDialog, True, True)
a.Write (iCalString)
a.Close

End Sub

Function createDay(rowNumber As Integer, dayNumber As Integer) As String

Dim iCalString As String

Dim subject As String
Dim teacher As String
Dim location As String
Dim description As String

Dim dayEvent As Date

Dim startEvent As Date
Dim startEventString As String    'in format YYYYMMDDTHHMMSS
Dim timeStart As Double
Dim timeStartString As String

Dim endEvent As Date
Dim endEventString As String      'in format YYYYMMDDTHHMMSS
Dim timeEnd As Double
Dim timeEndString As String

Dim timestampEvent As String 'in format YYYYMMDDTHHMMSS

timestampEvent = Format(CDate(lastUpdate), "yyyymmdd") & "T000000Z"

'lessons for HHMMSS
' 1    lesson    08:00   -   08:45
' 2    lesson    08:45   -   09:30
' break 15 min
' 3    lesson    09:45   -   10:30
' 4    lesson    10:30   -   11:15
' break 10 min
' 5    lesson    11:25   -   12:10
' break 45 min
' 6    lesson    12:55   -   13:40
' 7    lesson    13:40   -   14:25
' break 10 min
' 8    lesson    14:35   -   15:20
' break 10 min
' 9    lesson    15:30   -   16:15
' 10   lesson    16:15   -   17:00

'Columns C to L -> Lesson 1 to 10
For columnNumber = 3 To 12
    'set vars back
    description = ""
    subject = ""
    teacher = ""
    location = ""
    'date as YYYYMMDDT
    dayEvent = CDate(startDate) + dayNumber
    
    startEvent = dayEvent 'have to add time
    startEventString = Format(dayEvent, "yyyymmdd") & "T"
    
    endEvent = dayEvent 'have to add time
    endEventString = Format(dayEvent, "yyyymmdd") & "T"

    
    'set time (-2 hours for GTC Timezone)
    Select Case columnNumber
        Case 3
            '1 lesson 08:00-08:45
            timeStartString = "060000"
            timeEndString = "064500"
            timeStart = TimeValue("06:00")
            timeEnd = TimeValue("06:45")
            
        Case 4
            '2 lesson 08:45-09:30
            timeStartString = "064500"
            timeEndString = "073000"
            timeStart = TimeValue("06:45")
            timeEnd = TimeValue("07:30")
            
        Case 5
            '3 lesson 09:45-10:30
            timeStartString = "074500"
            timeEndString = "083000"
            timeStart = TimeValue("07:45")
            timeEnd = TimeValue("08:30")
            
        Case 6
            '4 lesson 10:30-11:15
            timeStartString = "083000"
            timeEndString = "091500"
            timeStart = TimeValue("08:30")
            timeEnd = TimeValue("09:15")
            
        Case 7
            '5 lesson 11:25-12:10
            timeStartString = "092500"
            timeEndString = "101000"
            timeStart = TimeValue("09:25")
            timeEnd = TimeValue("10:10")
            
        Case 8
            '6 lesson 12:55-13:40
            timeStartString = "105500"
            timeEndString = "114000"
            timeStart = TimeValue("10:55")
            timeEnd = TimeValue("11:40")
            
        Case 9
            '7 lesson 13:50-14:35
            timeStartString = "135000"
            timeEndString = "143500"
            timeStart = TimeValue("13:50")
            timeEnd = TimeValue("14:35")
            
        Case 10
            '8 lesson 14:35-15:20
            timeStartString = "123500"
            timeEndString = "132000"
            timeStart = TimeValue("12:35")
            timeEnd = TimeValue("13:20")
            
        Case 11
            '11 lesson 15:30-16:15
            timeStartString = "133000"
            timeEndString = "141500"
            timeStart = TimeValue("13:30")
            timeEnd = TimeValue("14:15")
            
        Case 12
            '12 lesson 16:15-17:00
            timeStartString = "141500"
            timeEndString = "150000"
            timeStart = TimeValue("14:15")
            timeEnd = TimeValue("15:00")
            
        Case Else
            'Error and end script
            MsgBox ("Error by Row " & rowNumber & ", Col. " & columnNumber & " | can't set time")
            Stop
    End Select
        
    'combine times and dates
    startEventString = startEventString & timeStartString & "Z"
    endEventString = endEventString & timeEndString & "Z"
        
    'get cell values or lessons
    If Left(Cells(rowNumber, columnNumber).Text, 1) = Chr(35) Then
        'for problems with Chr(35) = # -> #NV
        subject = "NV"
    Else
        subject = Cells(rowNumber, columnNumber).Text
        teacher = Cells(rowNumber + 1, columnNumber).Text
        
        
        'color cells to debug code
        Dim RandomR As Integer
        Dim RandomG As Integer
        Dim RandomB As Integer
        
        RandomR = Int((250 - 0 + 1) * Rnd + 0)
        RandomG = Int((250 - 0 + 1) * Rnd + 0)
        RandomB = Int((250 - 0 + 1) * Rnd + 0)
        Cells(rowNumber, columnNumber).Interior.Color = RGB(RandomR, RandomG, RandomB)
    End If
    'MsgBox ("Zeile: " & rowNumber & " Spalte: " & columnNumber & " Wert: " & subject)
    
    'get eventö-.lQ,kmjn vcx
    Select Case subject
        Case ""
        Case "NV"
        Case "#NV"
        Case "Betrieb"
        Case "Betrieb/Ferien"
        Case "Feiertag"
        
        Case Else
            'lessons at school
            iCalString = iCalString & "BEGIN:VEVENT" & insertNewLine
            
            iCalString = iCalString & "DTSTART;TZID=Europe/Berlin:" & startEventString & insertNewLine
            iCalString = iCalString & "DTEND;TZID=Europe/Berlin:" & endEventString & insertNewLine
            iCalString = iCalString & "DTSTAMP;TZID=Europe/Berlin:" & Format(CDate(lastUpdate), "yyyymmdd") & "T000000Z" & insertNewLine
            
            'title
            iCalString = iCalString & "SUMMARY:" & subject & insertNewLine
    
            'description
            If Cells(rowNumber, columnNumber) = "Sport" Then
                description = description & "Bitte beachten Sie die Aushänge bezüglich des Sportunterrichts" & "\n\n"
            End If
            If teacher <> "" Then
                description = description & subject & " mit " & teacher & "\n"
            Else
                description = description & subject & "\n"
            End If
            description = description & "Stand:" & lastUpdate & " für " & schoolClassName
            iCalString = iCalString & "DESCRIPTION:" & description & insertNewLine
            
            'unique ID from Update+StartDateTime+EndDateTime in UNIX/Integer
            iCalString = iCalString & "UID:AFBB-" & getSchoolClassName
                iCalString = iCalString & "-" & Round((lastUpdate - DateSerial(1970, 1, 1) + 0) * 86400)
                iCalString = iCalString & "-" & Round((startEvent - DateSerial(1970, 1, 1) + timeStart) * 86400)
                iCalString = iCalString & "-" & Round((endEvent - DateSerial(1970, 1, 1) + timeEnd) * 86400)
            iCalString = iCalString & insertNewLine
            
            'location in words
            If Cells(rowNumber, columnNumber) = "Sport" Then
                location = "XXL Dresden / AFBB Dresden"
            Else
                location = "AFBB Dresden"
            End If
            iCalString = iCalString & "LOCATION:" & location & insertNewLine
            
            'url
            iCalString = iCalString & "URL;VALUE=URI:https://www.afbb.de/de/dresden/intern_dresden.html" & insertNewLine
            
            'end event
            iCalString = iCalString & "END:VEVENT" & insertNewLine
        End Select
            
            
Next columnNumber

createDay = iCalString

End Function

Function insertNewLine() As String

'new line and object in ics-file
'insertNewLine = "\r\n"
insertNewLine = Chr(10)

End Function

Function getSchoolClassName() As String

getSchoolClassName = Range("A4").Value

End Function

Function getLastUpdate() As Date

updateString = Range("A5")
updateString = Replace(updateString, "Stand: ", "")
getLastUpdate = CDate(updateString)

End Function

Function getStartYear() As String

yearsString = Range("A2").Value
getStartYear = Left(yearsString, 4)

End Function

Function getEndYear() As String

yearsString = Range("A2").Value
getEndYear = Right(yearsString, 4)

End Function

Function getStartDate() As Date

startDateString = Range("A6").Value
getStartDate = CDate(startDateString)

End Function

Function getEndDate() As Date

Dim endDate As Date
Dim startDateCDate As Date

Dim rowNumber As Integer
rowNumber = 6
'Start by cell A6 as start date

Dim endFound As Boolean
endFound = False


Do While endFound = False
    If Range("A" & CStr(rowNumber)).Interior.Color <> "10079487" Then
        endFound = True
        rowNumber = rowNumber - 6
    Else
        rowNumber = rowNumber + 5 'step 5, because of 5 workday weeks
    End If
Loop


endDateString = Range("A" & CStr(rowNumber)).Value
endDate = CDate(endDateString)
startDateCDate = startDate

'when end date before start date then change end date to next year
' e.g. 24.08.2019 and 17.07.2019 change to 24.08.2019 and 17.07.2020
If startDate >= endDate Then
    endDate = DateAdd("yyyy", 1, endDate)
End If

getEndDate = endDate

End Function

Function worksheetExists(shtName As String, Optional wb As Workbook) As Boolean
    Dim sht As Worksheet

    If wb Is Nothing Then Set wb = ThisWorkbook
    On Error Resume Next
    Set sht = wb.Sheets(shtName)
    On Error GoTo 0
    worksheetExists = Not sht Is Nothing
End Function

