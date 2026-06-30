Option Explicit

Public Sub Main()
    GetDetails
End Sub

Private Sub GetDetails()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Dim cellText As String
    Dim employeeName As String
    Dim weekEnding As String

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

    Set ws = ActiveSheet

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
                weekEnding = Format( _
                    ws.Cells(i, "A").Value, _
                    "yyyy-mm-dd" _
                )
            End If

            waitingForWeekEnding = False

        ElseIf StrComp(cellText, "Action menu", vbTextCompare) = 0 Then

            If foundHours Then

                PrintProject _
                    employeeName, _
                    weekEnding, _
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

            totalHours = CDbl(cellText)
            foundHours = True

        ElseIf InStr(cellText, "・) > 0 Then

            dashPosition = InStr(cellText, "・)

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

Private Sub PrintProject( _
    ByVal employeeName As String, _
    ByVal weekEnding As String, _
    ByVal projectCodeA As String, _
    ByVal projectCodeB As String, _
    ByVal projectName As String, _
    ByVal totalHours As Double _
)

    Debug.Print "Name: " & employeeName
    Debug.Print "Week ending: " & weekEnding
    Debug.Print "Project Code A: " & projectCodeA
    Debug.Print "Project Code B: " & projectCodeB
    Debug.Print "Project Name: " & projectName
    Debug.Print "Hours: " & totalHours
    Debug.Print "----------------------"

End Sub

