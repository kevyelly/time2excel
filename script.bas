Option Explicit

' --------------------------------------------------------------------------
' Aggregation stores (module level, populated by the parser, consumed by the
' writer). Scripting.Dictionary preserves insertion order, which we rely on to
' keep project groups and people in first-appearance order.
'
'   mGroups : projectName -> persons dictionary
'   persons : name        -> rowInfo dictionary  ("code" + "hrs")
'   rowInfo : "code" -> "CODE_A / CODE_B"
'             "hrs"  -> hours dictionary  (weekSerialKey -> Double)
'   mWeeks  : weekSerialKey -> week serial (Double)
' --------------------------------------------------------------------------
Private mGroups As Object
Private mWeeks As Object

Public Sub Main()
    GenerateSummary
End Sub

Private Sub GenerateSummary()

    Dim source As Worksheet

    Set mGroups = CreateObject("Scripting.Dictionary")
    Set mWeeks = CreateObject("Scripting.Dictionary")

    ' Read the dump by name so re-runs work even when a SUMMARY sheet is active;
    ' fall back to the active sheet if there is no "Raw Data" tab.
    On Error Resume Next
    Set source = ThisWorkbook.Worksheets("Raw Data")
    On Error GoTo 0
    If source Is Nothing Then Set source = ActiveSheet

    ParseRawData source
    WriteSummarySheet

End Sub

Private Sub ParseRawData(ByVal ws As Worksheet)

    Dim lastRow As Long
    Dim i As Long

    Dim cellText As String
    Dim employeeName As String
    Dim weekSerial As Double

    Dim projectCodeA As String
    Dim projectCodeB As String
    Dim projectName As String
    Dim totalHours As Double

    Dim markerPosition As Long
    Dim dashPosition As Long

    Dim leftPart As String
    Dim rightPart As String

    Dim waitingForWeekEnding As Boolean
    Dim foundHours As Boolean

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For i = 1 To lastRow

        cellText = Trim(CStr(ws.Cells(i, "A").Value))

        If InStr(1, cellText, "App status:", vbTextCompare) > 0 Then

            markerPosition = InStr( _
                1, _
                cellText, _
                "Profile photo", _
                vbTextCompare _
            )

            employeeName = Trim(Mid( _
                cellText, _
                markerPosition + Len("Profile photo") _
            ))

            projectCodeA = ""
            projectCodeB = ""
            projectName = ""
            totalHours = 0
            foundHours = False

        ElseIf StrComp(cellText, "Week ending:", vbTextCompare) = 0 Then

            waitingForWeekEnding = True

        ElseIf waitingForWeekEnding And cellText <> "" Then

            If IsDate(ws.Cells(i, "A").Value) Then
                ' Value2 keeps the raw Excel serial (e.g. 46171), which is what
                ' we use as the week key and column header.
                weekSerial = ws.Cells(i, "A").Value2
            End If

            waitingForWeekEnding = False

        ElseIf StrComp(cellText, "Action menu", vbTextCompare) = 0 Then

            If foundHours Then

                StoreProject _
                    employeeName, _
                    weekSerial, _
                    projectCodeA, _
                    projectCodeB, _
                    projectName, _
                    totalHours

                projectCodeA = ""
                projectCodeB = ""
                projectName = ""
                totalHours = 0
                foundHours = False

            End If

        ElseIf IsNumeric(cellText) And projectCodeB <> "" Then

            ' Daily hours and the repeated total all land here; the final value
            ' is the week total, which is what we keep.
            totalHours = CDbl(cellText)
            foundHours = True

        ElseIf InStr(cellText, ChrW(8211)) > 0 Then

            dashPosition = InStr(cellText, ChrW(8211))

            leftPart = Trim(Left(cellText, dashPosition - 1))
            rightPart = Trim(Mid(cellText, dashPosition + 1))

            If projectCodeA = "" Then

                projectCodeA = leftPart

            ElseIf projectCodeB = "" Then

                projectCodeB = leftPart
                projectName = rightPart

            End If

        End If

    Next i

End Sub

Private Sub StoreProject( _
    ByVal employeeName As String, _
    ByVal weekSerial As Double, _
    ByVal projectCodeA As String, _
    ByVal projectCodeB As String, _
    ByVal projectName As String, _
    ByVal totalHours As Double _
)

    Dim projectCode As String
    Dim personProjectKey As String
    Dim weekKey As String

    Dim persons As Object
    Dim rowInfo As Object
    Dim hrs As Object

    projectName = CleanProjectName(projectName)

    projectCode = projectCodeA & " / " & projectCodeB

    personProjectKey = employeeName & Chr(30) & _
                       projectCodeA & Chr(30) & _
                       projectCodeB

    weekKey = CStr(weekSerial)

    ' Track every week found for the summary columns.
    mWeeks(weekKey) = weekSerial

    ' Create the project-name group when it does not exist yet.
    If Not mGroups.Exists(projectName) Then
        mGroups.Add projectName, CreateObject("Scripting.Dictionary")
    End If

    Set persons = mGroups(projectName)

    ' Create a separate row for every unique:
    ' employee + code A + code B combination.
    If Not persons.Exists(personProjectKey) Then

        Set rowInfo = CreateObject("Scripting.Dictionary")

        rowInfo("name") = employeeName
        rowInfo("codeA") = projectCodeA
        rowInfo("codeB") = projectCodeB
        rowInfo("code") = projectCode

        Set rowInfo("hrs") = CreateObject("Scripting.Dictionary")

        persons.Add personProjectKey, rowInfo

    End If

    Set rowInfo = persons(personProjectKey)
    Set hrs = rowInfo("hrs")

    ' Add hours when the exact same employee, project,
    ' code pair and week appears more than once.
    If hrs.Exists(weekKey) Then
        hrs(weekKey) = CDbl(hrs(weekKey)) + totalHours
    Else
        hrs.Add weekKey, totalHours
    End If

End Sub

' The template drops a leading "PH " that the source name sometimes carries
' (e.g. "PH Duskin AMS" -> "Duskin AMS"). Flip this to keep names verbatim.
Private Function CleanProjectName(ByVal projectName As String) As String

    If Left(projectName, 3) = "PH " Then
        CleanProjectName = Trim(Mid(projectName, 4))
    Else
        CleanProjectName = projectName
    End If

End Function

Private Sub WriteSummarySheet()

    Dim ws As Worksheet
    Dim sheetName As String

    Dim weekSerials() As Double
    Dim numWeeks As Long

    Dim numGroups As Long
    Dim numPersons As Long

    Dim totalRows As Long
    Dim numCols As Long

    Dim data() As Variant

    Dim projKey As Variant
    Dim personKey As Variant
    Dim persons As Object
    Dim rowInfo As Object
    Dim hrs As Object

    Dim r As Long
    Dim j As Long
    Dim g As Long
    Dim weekKey As String

    ' ---- Column order: unique week serials, chronologically ascending -------
    weekSerials = SortedWeekSerials()
    numWeeks = mWeeks.Count

    ' ---- Sheet geometry -----------------------------------------------------
    numGroups = mGroups.Count
    numPersons = 0
    For Each projKey In mGroups.Keys
        numPersons = numPersons + mGroups(projKey).Count
    Next projKey

    ' header + all person rows + one blank row between consecutive groups
    totalRows = 1 + numPersons
    If numGroups > 1 Then totalRows = totalRows + (numGroups - 1)
    numCols = 3 + numWeeks

    ReDim data(1 To totalRows, 1 To numCols)

    ' ---- Header row: labels in A:C, week serials from column D onward -------
    data(1, 1) = "Project Name"
    data(1, 2) = "Name"
    data(1, 3) = "Project Code"
    For j = 1 To numWeeks
        data(1, 3 + j) = weekSerials(j)
    Next j

    ' ---- Body: grouped by project, one blank spacer row between groups ------
    r = 1
    g = 0
    For Each projKey In mGroups.Keys

        g = g + 1
        Set persons = mGroups(projKey)

        For Each personKey In persons.Keys

            Set rowInfo = persons(personKey)
            Set hrs = rowInfo("hrs")

            r = r + 1
            data(r, 1) = projKey
            data(r, 2) = rowInfo("name")
            data(r, 3) = rowInfo("code")

            For j = 1 To numWeeks
                weekKey = CStr(weekSerials(j))
                If hrs.Exists(weekKey) Then
                    data(r, 3 + j) = hrs(weekKey)
                End If
            Next j

        Next personKey

        ' blank spacer row between groups (not after the last one)
        If g < numGroups Then r = r + 1

    Next projKey

    ' ---- Create / replace the dated Summary sheet --------------------------
    sheetName = "SUMMARY " & Format(Date, "yyyy-mm-dd")

    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(sheetName).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set ws = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = sheetName

    ' ---- Dump in one shot, then beautify -----------------------------------
    ws.Range(ws.Cells(1, 1), ws.Cells(totalRows, numCols)).Value = data

    BeautifySheet ws, totalRows, numCols, numWeeks

End Sub

' Apply colours, borders and layout to the freshly written Summary sheet.
Private Sub BeautifySheet( _
    ByVal ws As Worksheet, _
    ByVal totalRows As Long, _
    ByVal numCols As Long, _
    ByVal numWeeks As Long _
)

    Dim r As Long
    Dim groupIdx As Long
    Dim prevBlank As Boolean
    Dim idFill As Long

    ' ---- Header bar: dark blue fill, white bold, centred -------------------
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, numCols))
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(1).RowHeight = 22

    If numWeeks > 0 Then
        ws.Range(ws.Cells(1, 4), ws.Cells(1, 3 + numWeeks)).NumberFormat = "m/d/yyyy"
    End If
    ApplyThinBorders ws.Range(ws.Cells(1, 1), ws.Cells(1, numCols))

    ' ---- Body: banded fill per project group, borders, centred hours ------
    groupIdx = 0
    prevBlank = True
    For r = 2 To totalRows

        If Len(CStr(ws.Cells(r, 1).Value) & CStr(ws.Cells(r, 2).Value)) = 0 Then

            ' spacer row between groups -> left plain
            prevBlank = True

        Else

            If prevBlank Then groupIdx = groupIdx + 1
            prevBlank = False

            ' alternate the identity-column shade group by group
            If groupIdx Mod 2 = 1 Then
                idFill = RGB(217, 225, 242)   ' light blue
            Else
                idFill = RGB(226, 239, 218)   ' light green
            End If

            ws.Range(ws.Cells(r, 1), ws.Cells(r, 3)).Interior.Color = idFill
            ws.Cells(r, 1).Font.Bold = True

            If numWeeks > 0 Then
                ws.Range(ws.Cells(r, 4), ws.Cells(r, numCols)).HorizontalAlignment = xlCenter
            End If

            ApplyThinBorders ws.Range(ws.Cells(r, 1), ws.Cells(r, numCols))

        End If

    Next r

    ' ---- Layout: widths, frozen header + identity columns, no gridlines ----
    ws.Columns.AutoFit
    If ws.Columns(1).ColumnWidth > 40 Then ws.Columns(1).ColumnWidth = 40

    ws.Activate
    With ActiveWindow
        .DisplayGridlines = False
        .FreezePanes = False
        .SplitRow = 1
        .SplitColumn = 3
        .FreezePanes = True
    End With

End Sub

' Light-grey thin border around every cell of the range.
Private Sub ApplyThinBorders(ByVal rng As Range)

    With rng.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(191, 191, 191)
    End With

End Sub

' Return the unique week serials as a 1-based array, sorted ascending.
Private Function SortedWeekSerials() As Double()

    Dim arr() As Double
    Dim n As Long
    Dim k As Variant
    Dim i As Long
    Dim jj As Long
    Dim tmp As Double

    n = mWeeks.Count
    ReDim arr(1 To Application.Max(n, 1))

    i = 0
    For Each k In mWeeks.Keys
        i = i + 1
        arr(i) = mWeeks(k)
    Next k

    ' insertion sort, ascending (chronological)
    For i = 2 To n
        tmp = arr(i)
        jj = i - 1
        Do While jj >= 1
            If arr(jj) <= tmp Then Exit Do
            arr(jj + 1) = arr(jj)
            jj = jj - 1
        Loop
        arr(jj + 1) = tmp
    Next i

    SortedWeekSerials = arr

End Function


