Attribute VB_Name = "Reports"
Option Explicit

Private Const msMODULE As String = "Reports"

Private rsStaReport As ADODB.Recordset
Private rsLineReport As ADODB.Recordset
Private wsStaReport As Worksheet
Private wsLineReport As Worksheet
Private cnSqlServer As SQLServerConn

Private msReport As String
Private msShift As String
Private msLine As String
Private msStation As String
Private msStartDate As String
Private msStartTime As String
Private msEndDate As String
Private msEndTime As String


'-- STATION REPORT --
'=====================
Sub Make_StationReport()
'Error handling
Const sSOURCE As String = "Make_StationReport()"
Const bDebug As Boolean = False
On Error GoTo ErrorHandler
Application.ScreenUpdating = False

'Allow user to halt program
Application.EnableCancelKey = xlErrorHandler

'Declare
Dim i As Integer
Dim arStations() As String, sSql As String

'Build sql statement
sSql = ""
If GetStation() = "All" Then
    sSql = Sql.StationReport & " ORDER BY WorkCenterObjId,Station"
Else
    arStations() = Split(msStation, ",")
    sSql = " WHERE Station IN ("
    For i = LBound(arStations) To UBound(arStations)
        sSql = sSql & " '" & arStations(i) & "',"
    Next i
    sSql = Left(sSql, Len(sSql) - 1)
    sSql = Sql.StationReport & " " & sSql & " ) ORDER BY WorkCenterObjId,Station"
End If

'/* DEBUG */
If Not bDebug Then
    Debug.Print sSql
    GoTo ErrorExit
End If

'SQL Server connection
Set cnSqlServer = New SQLServerConn
cnSqlServer.ConnOpen

'Retrieve data
Set rsStaReport = New ADODB.Recordset
rsStaReport.Open sSql, cnSqlServer.Conn, adOpenStatic

'Create new worksheet and format with report template
If Not Template_StationReport() Then Err.Raise glHANDLED_ERROR

'Populate report template with query data
If Not rsStaReport.BOF And Not rsStaReport.EOF Then
    'Copy data to worksheet
    wsStaReport.Range("B5").CopyFromRecordset rsStaReport
    
    'Format output worksheet
    If Not Format_StationReport() Then Err.Raise glHANDLED_ERROR
Else
    With wsStaReport.Range("B5:J5")
        .Merge
        .HorizontalAlignment = xlCenter
        .Value = " NO RECORDS FOUND .... "
        .Font.Bold = True
    End With
End If

ErrorExit:
'Clear memory
On Error Resume Next
Application.ScreenUpdating = True
cnSqlServer.ConnClose
rsStaReport.Close
Set rsStaReport = Nothing
Exit Sub

ErrorHandler:
If bCentralErrorHandler(msMODULE, sSOURCE, , True) Then
    Stop
    Resume
Else
    Resume ErrorExit
End If
End Sub

'-- STATION -> LAST 5 SHIFT --
'==============================
Sub Make_StationFiveReport()
'Error handling
Const sSOURCE As String = "Make_StationFiveReport()"
Const bDebug As Boolean = False
On Error GoTo ErrorHandler
Application.ScreenUpdating = False

'Declare
Dim i As Integer, lRow As Long
Dim arStations() As String, sSql As String
Dim sStartDate As String, sEndDate As String
Dim sStartTime As String, sEndTime As String
Dim sShift As String, sPeriodEnd As String

'Calculate 5 production working days
If Format(Now(), "ddd") = "Mon" Then
    sEndDate = Format(Now() - 3, "yyy-mm-dd")
ElseIf Format(Now(), "ddd") = "Sun" Then
    sEndDate = Format(Now() - 2, "yyyy-mm-dd")
Else
    sEndDate = Format(Now() - 1, "yyyy-mm-dd")
End If
For i = 0 To 5
    sStartDate = Format(Now() - (5 + i), "yyyy-mm-dd")
    If Application.WorksheetFunction.NetworkDays(sStartDate, sEndDate) = 5 Then Exit For
Next i

'Define start & end times
If GetShift() = "0" Then
    sEndDate = Format(Format(sEndDate, "yyyy/mm") & "/" & Format(sEndDate, "dd") + 1, "yyyy/mm/dd")
    sStartTime = "06:45"
    sEndTime = "06:45"
    sShift = "%"
ElseIf GetShift() = "1" Then
    sStartTime = "06:45"
    sEndTime = "14:45"
    sShift = "A"
ElseIf GetShift() = "2" Then
    sStartTime = "14:45"
    sEndTime = "22:45"
    sShift = "B"
ElseIf GetShift() = "3" Then
    sEndDate = Format(Format(sEndDate, "yyyy/mm") & "/" & Format(sEndDate, "dd") + 1, "yyyy/mm/dd")
    sStartTime = "22:45"
    sEndTime = "06:45"
    sShift = "C"
End If

'Build sql statement
sSql = ""
If GetStation() = "All" Then
    sSql = Sql.StationFiveShift(sStartDate, sEndDate, sShift) & " ORDER BY st.WorkCenterObjId, t.Station, t.ProdDate "
Else
    arStations() = Split(GetStation(), ",")
    sSql = " WHERE Station IN ("
    For i = LBound(arStations) To UBound(arStations)
        sSql = sSql & " '" & arStations(i) & "',"
    Next i
    sSql = Left(sSql, Len(sSql) - 1)
    sSql = Sql.StationFiveShift(sStartDate, sEndDate, sShift) & " " & sSql & " ) ORDER BY st.WorkCenterObjId, t.Station, t.ProdDate "
End If

'/* DEBUG */
If bDebug Then
    Debug.Print sSql
    GoTo ErrorExit
End If

'SQLServer connection
Set cnSqlServer = New SQLServerConn
cnSqlServer.ConnOpen

'Retrieve data
Set rsStaReport = New ADODB.Recordset
rsStaReport.Open sSql, cnSqlServer.Conn, adOpenStatic

'Create new worksheet and format with report template
If Not Template_StationReport() Then Err.Raise glHANDLED_ERROR

If Not rsStaReport.BOF And Not rsStaReport.EOF Then
    'Populate report template with extracted data
    wsStaReport.Range("B5").CopyFromRecordset rsStaReport
    
    'Update PeriodStart & PeriodEnd cells
    For lRow = 0 To rsStaReport.RecordCount - 1
        With wsStaReport
            'Define PeriodEnd field, change to next day if All shift or 3rd shift
            If GetShift() = "0" Or GetShift() = "3" Then
                 sPeriodEnd = Format(Format(wsStaReport.Range("C" & (lRow + 5)).Value, "yyyy/mm") & "/" & Format(wsStaReport.Cells(lRow + 5, 3).Value, "dd") + 1, "yyyy/mm/dd")
            Else
                sPeriodEnd = wsStaReport.Range("C" & (lRow + 5)).Value
            End If
            sPeriodEnd = Format(sPeriodEnd & " " & sEndTime, "yyyy-mm-dd hh:mm")
            .Range("D" & (lRow + 5)).Value = sPeriodEnd
            
            'Define PeriodStart field
            .Range("C" & (lRow + 5)).Value = Format(wsStaReport.Range("C" & (lRow + 5)).Value & " " & sStartTime, "yyyy-mm-dd hh:mm")
        End With
    Next lRow
    
    'Format report output
    If Not Format_StationFiveShift() Then Err.Raise glHANDLED_ERROR
Else
    With wsStaReport.Range("B5:N5")
        .Merge
        .HorizontalAlignment = xlCenter
        .Value = " NO RECORDS FOUND .... "
        .Font.Bold = True
    End With
End If
        
'Clear memory
ErrorExit:
On Error Resume Next
Application.ScreenUpdating = True
cnSqlServer.ConnClose
If rsStaReport Is Nothing Then rsStaReport.Close
Set rsStaReport = Nothing
Exit Sub

ErrorHandler:
If bCentralErrorHandler(msMODULE, sSOURCE, , True) Then
    Stop
    Resume
Else
    Resume ErrorExit
End If
End Sub

Private Function Format_StationReport() As Boolean
'Error handling
Const sSOURCE As String = "Format_StationReport()"
On Error GoTo ErrorHandler
Format_StationReport = True

'Declare
Dim lLastRow As Long, lRow As Long
lLastRow = wsStaReport.Range("B4").End(xlDown).Row

'Start & End Period - Indent right alignment
With wsStaReport.Range("C5:D" & lLastRow)
    .HorizontalAlignment = xlRight
    .IndentLevel = 1
End With

'Total values columns - Right alignment
With wsStaReport.Range("E5:N" & lLastRow)
    .HorizontalAlignment = xlRight
    .IndentLevel = 1
End With

'Freeze pane
wsStaReport.Range("C5").Select
ActiveWindow.FreezePanes = True

wsStaReport.Range("C5:D" & lLastRow).NumberFormat = "yyyy/mm/dd hh:mm:ss"

For lRow = 5 To lLastRow

    'Alternate row highlight
    If Not lRow Mod 2 = 0 Then
        wsStaReport.Range("B" & lRow & ":N" & lRow).Interior.Color = RGB(230, 230, 230)
    End If
    
    'Calculate Pass Ratio & format with GA color code
    '  Green>=99, Yellow>=97 & <99, Red<97
    With wsStaReport.Cells(lRow, 8)
        If .Value >= 99 Then
            .Interior.ColorIndex = 43               'Highlight green
            .Font.ColorIndex = 10                   'Font color dark green
        ElseIf .Value >= 97 And .Value < 99 Then
            .Interior.ColorIndex = 36               'Highlight yellow
            .Font.ColorIndex = 12                   'Font color dark yellow
        ElseIf .Value < 97 And .Value >= 1 Then
            .Interior.ColorIndex = 46               'Highlight light red
            .Font.ColorIndex = 9                    'Font color dark red
        ElseIf .Value = 0 Then
            'alternate color if ratio=0
            If lRow Mod 2 = 0 Then
                .Interior.ColorIndex = xlNone       'Highlight no color
            Else
                .Interior.Color = RGB(230, 230, 230)    'Highlight light gray
            End If
            .Font.ColorIndex = 9                        'Font color dark red
        End If
        'Add percent sign to ratio value
        .Value = wsStaReport.Range("H" & lRow).Value & " %"
    End With
    
    'Calculate Scan Ratio & format with GA color code
    '  Green>=95, Yellow>=90 & <95, Red<90
    With wsStaReport.Cells(lRow, 9)
        If .Value >= 95 Then
            .Interior.ColorIndex = 43               'Highlight green
            .Font.ColorIndex = 10                   'Font color dark green
        ElseIf .Value >= 90 And .Value < 95 Then
            .Interior.ColorIndex = 36               'Highlight yellow
            .Font.ColorIndex = 12                   'Font color dark yellow
        ElseIf .Value < 90 And .Value >= 1 Then
            .Interior.ColorIndex = 46               'Highlight light red
            .Font.ColorIndex = 9                    'Font color dark red
        ElseIf .Value = 0 Then
            'alternate color if ratio=0
            If lRow Mod 2 = 0 Then
                .Interior.ColorIndex = xlNone       'Highlight no color
            Else
                .Interior.Color = RGB(230, 230, 230)    'Highlight light gray
            End If
            .Font.ColorIndex = 9                        'Font color dark red
        End If
        'Add percent sign to ratio value
        .Value = wsStaReport.Range("I" & lRow).Value & " %"
    End With
    
Next lRow

'Auto fit rows & columns
wsStaReport.Rows("4:4").RowHeight = 28
wsStaReport.Columns.AutoFit

ErrorExit:
On Error Resume Next
Exit Function

ErrorHandler:
Format_StationReport = False
If bCentralErrorHandler(msMODULE, sSOURCE) Then
    Stop
    Resume
Else
    Resume ErrorExit
End If
End Function

Private Function Format_StationFiveShift() As Boolean
'Error handling
Const sSOURCE As String = "Format_StationFiveShift()"
On Error GoTo ErrorHandler
Format_StationFiveShift = True

'Declare
Dim lLastRow As Long, lRow As Long, lCol As Long
Dim sPrev As String, lStationCnt As Long
lLastRow = wsStaReport.Range("B4").End(xlDown).Row

'Start & End Period - Center alignment
With wsStaReport.Range("C5:D" & lLastRow)
    .HorizontalAlignment = xlRight
    .IndentLevel = 1
End With

'Total values columns - Right alignment
With wsStaReport.Range("E5:N" & lLastRow)
    .HorizontalAlignment = xlRight
    .IndentLevel = 1
End With

'Freeze pane
wsStaReport.Range("C5").Select
ActiveWindow.FreezePanes = True

wsStaReport.Range("C5:D" & lLastRow).NumberFormat = "yyyy/mm/dd hh:mm:ss"

'Alternate highlighting color
sPrev = wsStaReport.Range("B5").Value
lStationCnt = 1
For lRow = 5 To lLastRow

    'Alternating highligh per station
    If Not wsStaReport.Range("B" & lRow).Value = sPrev Then
        lStationCnt = lStationCnt + 1
        sPrev = wsStaReport.Range("B" & lRow).Value
    End If
    If Not lStationCnt Mod 2 = 0 Then
        wsStaReport.Range("B" & lRow & ":N" & lRow).Interior.Color = RGB(230, 230, 230)
    End If
    
    'Calculate Pass Ratio & format with GA color code
    '  Green>=99, Yellow>=97 & <99, Red<97
    With wsStaReport.Range("H" & lRow)
        If .Value >= 99 Then
            .Interior.ColorIndex = 43               'Highlight green
            .Font.ColorIndex = 10                   'Font color dark green
        ElseIf .Value >= 97 And .Value < 99 Then
            .Interior.ColorIndex = 36               'Highlight yellow
            .Font.ColorIndex = 12                   'Font color dark yellow
        ElseIf .Value < 97 And .Value >= 1 Then
            .Interior.ColorIndex = 46               'Highlight light red
            .Font.ColorIndex = 9                    'Font color dark red
        ElseIf .Value = 0 Then
            'alternate color if ratio=0
            If lStationCnt Mod 2 = 0 Then
                .Interior.ColorIndex = xlNone       'Highlight no color
            Else
                .Interior.Color = RGB(230, 230, 230)    'Highlight light gray
            End If
            .Font.ColorIndex = 9                        'Font color dark red
        End If
        'Add percent sign to ratio value
        .Value = wsStaReport.Range("H" & lRow).Value & " %"
    End With
    
    'Calculate Scan Ratio & format with GA color code
    '  Green>=95, Yellow>=90 & <95, Red<90
    With wsStaReport.Range("I" & lRow)
        If .Value >= 95 Then
            .Interior.ColorIndex = 43               'Highlight green
            .Font.ColorIndex = 10                   'Font color dark green
        ElseIf .Value >= 90 And .Value < 95 Then
            .Interior.ColorIndex = 36               'Highlight yellow
            .Font.ColorIndex = 12                   'Font color dark yellow
        ElseIf .Value < 90 And .Value >= 1 Then
            .Interior.ColorIndex = 46               'Highlight light red
            .Font.ColorIndex = 9                    'Font color dark red
        ElseIf .Value = 0 Then
            'alternate color if ratio=0
            If lStationCnt Mod 2 = 0 Then
                .Interior.ColorIndex = xlNone       'Highlight no color
            Else
                .Interior.Color = RGB(230, 230, 230)    'Highlight light gray
            End If
            .Font.ColorIndex = 9                        'Font color dark red
        End If
        'Add percent sign to ratio value
        .Value = wsStaReport.Range("I" & lRow).Value & " %"
    End With
        
Next lRow

'Auto fit rows & columns
wsStaReport.Rows("4:4").RowHeight = 28
wsStaReport.Columns.AutoFit

ErrorExit:
On Error Resume Next
Exit Function

ErrorHandler:
Format_StationFiveShift = False
If bCentralErrorHandler(msMODULE, sSOURCE) Then
    Stop
    Resume
Else
    Resume ErrorExit
End If
End Function

Private Function Template_StationReport() As Boolean
'Error handling
Const sSOURCE As String = "Template_StationReport()"
On Error GoTo ErrorHandler
Template_StationReport = True

'Declare
Dim iSheet As Integer
Dim ws As Worksheet
Dim sShiftTitle As String, sReportTitle As String
Dim sShtName As String

'Build report sheet name
Select Case GetShift()
    Case Is = "0"
        sShtName = "All Shifts"
    Case Is = "1"
        sShtName = "1st Shift"
    Case Is = "2"
        sShtName = "2nd Shift"
    Case Is = "3"
        sShtName = "3rd Shift"
End Select
Select Case GetReport()
    Case Is = "0"
        sShtName = "Single Sta Trend"
    Case Is = "1"
        sShtName = sShtName & " - Last Shift"
    Case Is = "2"
        sShtName = sShtName & " - Last 5 Shifts"
End Select

'Delete existing worksheet
For Each ws In Worksheets
    Application.DisplayAlerts = False
    If ws.Name = sShtName Then ws.Delete
Next ws

'Create new report worksheet
Set wsStaReport = Sheets.Add
wsStaReport.Name = sShtName
wsStaReport.Select
Application.DisplayAlerts = True

'Title1 / Plant Name
With wsStaReport.Range("B1:N1")
    .Merge
    .HorizontalAlignment = xlCenter
    .Font.Size = 12
    .Font.Bold = True
    .Value = "HMMA GA TDM"
End With
'Title 2 / Shift Description
With wsStaReport.Range("B2:N2")
    .Merge
    .HorizontalAlignment = xlCenter
    .Font.Size = 12
    .Font.Bold = True
    Select Case GetShift()
        Case Is = "0"
            sShiftTitle = "All Shifts - Station Level"
        Case Is = "1"
            sShiftTitle = "1st Shift - Station Level"
        Case Is = "2"
            sShiftTitle = "2nd Shift - Station Level"
        Case Is = "3"
            sShiftTitle = "3rd Shift - Station Level"
    End Select
    .Value = sShiftTitle
End With
'Title3 / Report Type Description
With wsStaReport.Range("B3:N3")
    .Merge
    .HorizontalAlignment = xlCenter
    .Font.Size = 12
    .Font.Bold = True
    Select Case GetReport()
        Case Is = "1"
            sReportTitle = "Last Completed Shift"
        Case Is = "2"
            sReportTitle = "Last 5 Completed Shifts"
        Case Is = "0"
            wsStaReport.Range("B2:N2").Value = "Station Level"
            sReportTitle = "Trending Period  " & GetStartDate() & " " & GetStartTime & " - " & GetEndDate & " " & GetEndTime
    End Select
    .Value = sReportTitle
End With

'Header row
With wsStaReport
    .Range("B4").Value = "Station Name"
    .Range("C4").Value = "Period Start"
    .Range("D4").Value = "Period End"
    .Range("E4").Value = "Total Vehicle"
    .Range("F4").Value = "Total Pass"
    .Range("G4").Value = "Total Scans"
    .Range("H4").Value = "Pass Ratio"
    .Range("I4").Value = "Scan Ratio"
    .Range("J4").Value = "Total NGs"
    .Range("K4").Value = "Torque High"
    .Range("L4").Value = "Torque Low"
    .Range("M4").Value = "Angle High"
    .Range("N4").Value = "Angle Low"
End With

'Format header row
With wsStaReport.Range("B4:N4")
    .Interior.ColorIndex = 41
    .Font.ColorIndex = 2
    .Font.Bold = True
    .AutoFilter
    .Columns.AutoFit
End With
wsStaReport.Rows("4:4").RowHeight = 28

ErrorExit:
On Error Resume Next
Set ws = Nothing
Exit Function

ErrorHandler:
Template_StationReport = False
If bCentralErrorHandler(msMODULE, sSOURCE) Then
    Stop
    Resume
Else
    Resume ErrorExit
End If
End Function

'-- LINE REPORT --
'==================
Sub Make_LineReport()
'Error handling
Const sSOURCE As String = "Make_LineReport()"
Const bDebug As Boolean = False
On Error GoTo ErrorHandler
Application.ScreenUpdating = False

'Declare
Dim sSql As String
Dim i As Integer, arLines() As String
              
'Build sql statement
sSql = ""
If GetLine() = "All" Then
    sSql = Sql.LineReport & " ORDER BY WorkCenterObjId"
Else
    arLines() = Split(GetLine(), ",")
    sSql = " WHERE WorkCenter IN ("
    For i = LBound(arLines) To UBound(arLines)
        sSql = sSql & " '" & arLines(i) & "',"
    Next i
    sSql = Left(sSql, Len(sSql) - 1)
    sSql = Sql.LineReport & " " & sSql & " ) ORDER BY WorkCenterObjId"
End If

'/* DEBUG */
If bDebug Then
    Debug.Print sSql
    GoTo ErrorExit
End If

'SQLServer connection
Set cnSqlServer = New SQLServerConn
cnSqlServer.ConnOpen

'Retrieve data
Set rsLineReport = New ADODB.Recordset
rsLineReport.Open sSql, cnSqlServer.Conn, adOpenStatic

'Create new worksheet and format with report template
If Not Template_LineReport() Then Err.Raise glHANDLED_ERROR

'Populate report template with query data
If Not rsLineReport.BOF And Not rsLineReport.EOF Then
    'Copy data to worksheet
    wsLineReport.Range("B5").CopyFromRecordset rsLineReport
    
    'Format output worksheet
    If Not Format_LineReport() Then Err.Raise glHANDLED_ERROR
Else
    'Indicate no records found with search criteria
    With wsLineReport.Range("B5:O5")
        .Merge
        .HorizontalAlignment = xlCenter
        .Value = " NO RECORDS FOUND .... "
        .Font.Bold = True
    End With
End If

ErrorExit:
'Clear memory
On Error Resume Next
Application.ScreenUpdating = True
cnSqlServer.ConnClose
rsLineReport.Close
Set rsLineReport = Nothing
Exit Sub

ErrorHandler:
If bCentralErrorHandler(msMODULE, sSOURCE, , True) Then
    Stop
    Resume
Else
    Resume ErrorExit
End If
End Sub

Sub Make_LineFiveShift()
'Error handling
Const sSOURCE As String = "Make_LineFiveShift()"
Const bDebug As Boolean = False
On Error GoTo ErrorHandler
Application.ScreenUpdating = False

'Declare
Dim sSql As String
Dim i As Integer, lRow As Long
Dim arLines() As String
Dim sStartDate As String, sEndDate As String
Dim sStartTime As String, sEndTime As String
Dim sShift As String, sPeriodEnd As String

'Calculate 5 production working days
If Format(Now(), "ddd") = "Mon" Then
    sEndDate = Format(Now() - 3, "yyy-mm-dd")
ElseIf Format(Now(), "ddd") = "Sun" Then
    sEndDate = Format(Now() - 2, "yyyy-mm-dd")
Else
    sEndDate = Format(Now() - 1, "yyyy-mm-dd")
End If
For i = 0 To 5
    sStartDate = Format(Now() - (5 + i), "yyyy-mm-dd")
    If Application.WorksheetFunction.NetworkDays(sStartDate, sEndDate) = 5 Then Exit For
Next i

'Define start & end times
If GetShift() = "0" Then
    sEndDate = Format(Format(sEndDate, "yyyy/mm") & "/" & Format(sEndDate, "dd") + 1, "yyyy/mm/dd")
    sStartTime = "06:45"
    sEndTime = "06:45"
    sShift = "%"
ElseIf GetShift() = "1" Then
    sStartTime = "06:45"
    sEndTime = "14:45"
    sShift = "A"
ElseIf GetShift() = "2" Then
    sStartTime = "14:45"
    sEndTime = "22:45"
    sShift = "B"
ElseIf GetShift() = "3" Then
    sEndDate = Format(Format(sEndDate, "yyyy/mm") & "/" & Format(sEndDate, "dd") + 1, "yyyy/mm/dd")
    sStartTime = "22:45"
    sEndTime = "06:45"
    sShift = "C"
End If

'Build sql statement
sSql = ""
If GetLine() = "All" Then
    sSql = Sql.LineFiveShift(sStartDate, sEndDate, sShift) & " ORDER BY WorkCenterObjId, PeriodStart "
Else
    arLines() = Split(GetLine(), ",")
    sSql = " WHERE WorkCenter IN ("
    For i = LBound(arLines) To UBound(arLines)
        sSql = sSql & " '" & arLines(i) & "',"
    Next i
    sSql = Left(sSql, Len(sSql) - 1)
    sSql = Sql.LineFiveShift(sStartDate, sEndDate, sShift) & " " & sSql & " ) ORDER BY WorkCenterObjId, PeriodStart "
End If

'/* DEBUG */
If bDebug Then
    Debug.Print sSql
    GoTo ErrorExit
End If

'SQLServer connection
Set cnSqlServer = New SQLServerConn
cnSqlServer.ConnOpen

'Retrieve data
Set rsLineReport = New ADODB.Recordset
rsLineReport.Open sSql, cnSqlServer.Conn, adOpenStatic

'Create new worksheet and format with report template
If Not Template_LineReport() Then Err.Raise glHANDLED_ERROR

If Not rsLineReport.BOF And Not rsLineReport.EOF Then
    'Populate report template with extracted data
    wsLineReport.Range("B5").CopyFromRecordset rsLineReport
    
    'Update PeriodStart & PeriodEnd cells
    For lRow = 0 To rsLineReport.RecordCount - 1
        With wsLineReport
            'Define PeriodEnd field, change to next day if All shift or 3rd shift
            If GetShift() = "0" Or GetShift() = "3" Then
                sPeriodEnd = Format(Format(.Range("C" & (lRow + 5)).Value, "yyyy/mm") & "/" & Format(.Range("C" & (lRow + 5)).Value, "dd") + 1, "yyyy/mm/dd")
            Else
                sPeriodEnd = .Range("C" & (lRow + 5)).Value
            End If
            sPeriodEnd = Format(sPeriodEnd & " " & sEndTime, "yyyy-mm-dd hh:mm")
            .Range("D" & (lRow + 5)).Value = sPeriodEnd
            
            'Define PeriodStart field
            .Range("C" & (lRow + 5)).Value = Format(.Range("C" & (lRow + 5)).Value & " " & sStartTime, "yyyy-mm-dd hh:mm")
        End With
    Next lRow
    
    'Format report output
    If Not Format_LineFiveShift() Then Err.Raise glHANDLED_ERROR
Else
    With wsLineReport.Range("B5:N5")
        .Merge
        .HorizontalAlignment = xlCenter
        .Value = " NO RECORDS FOUND .... "
        .Font.Bold = True
    End With
End If
        
'Clear memory
ErrorExit:
On Error Resume Next
Application.ScreenUpdating = True
cnSqlServer.ConnClose
rsStaReport.Close
Set rsStaReport = Nothing
Exit Sub

ErrorHandler:
If bCentralErrorHandler(msMODULE, sSOURCE, , True) Then
    Stop
    Resume
Else
    Resume ErrorExit
End If
End Sub

Private Function Format_LineReport() As Boolean
'Error handling
Const sSOURCE As String = "Format_LineReport()"
On Error GoTo ErrorHandler
Format_LineReport = True

'Declare
Dim lLastRow As Long, lRow As Long
Dim lCol As Long
lLastRow = wsLineReport.Range("B4").End(xlDown).Row

'Freeze pane
wsLineReport.Range("B5").Select
ActiveWindow.FreezePanes = True

'Right align values
With wsLineReport.Range("C5:O" & lLastRow)
    .HorizontalAlignment = xlRight
    .IndentLevel = 1
End With

'Period start & end date/time format
wsLineReport.Range("C5:D" & lLastRow).NumberFormat = "yyyy-mm-dd hh:mm"

For lRow = 5 To lLastRow
    'Alternate row highlight
    If Not lRow Mod 2 = 0 Then
        wsLineReport.Range("B" & lRow & ":O" & lRow).Interior.Color = RGB(230, 230, 230)
    End If
    
    'Calculate Avg Pass Ratio & format with GA color code
    '  Green>=99, Yellow>=97 & <99, Red<97
    With wsLineReport.Range("G" & lRow)
        If .Value >= 99 Then
            .Interior.ColorIndex = 43               'Highlight green
            .Font.ColorIndex = 10                   'Font color dark green
        ElseIf .Value >= 97 And .Value < 99 Then
            .Interior.ColorIndex = 36               'Highlight yellow
            .Font.ColorIndex = 12                   'Font color dark yellow
        ElseIf .Value < 97 And .Value >= 1 Then
            .Interior.ColorIndex = 46               'Highlight light red
            .Font.ColorIndex = 9                    'Font color dark red
        ElseIf .Value = 0 Then
            'alternate color if ratio=0
            If lRow Mod 2 = 0 Then
                .Interior.ColorIndex = xlNone       'Highlight no color
            Else
                .Interior.Color = RGB(230, 230, 230)    'Highlight light gray
            End If
            .Font.ColorIndex = 9                        'Font color dark red
        End If
        'Add percent sign to ratio value
        .Value = wsLineReport.Range("G" & lRow).Value & " %"
    End With
    
    'Calculate Avg Scan Ratio & format with GA color code
    '  Green>=95, Yellow>=90 & <95, Red<90
    With wsLineReport.Range("J" & lRow)
        If .Value >= 95 Then
            .Interior.ColorIndex = 43               'Highlight green
            .Font.ColorIndex = 10                   'Font color dark green
        ElseIf .Value >= 90 And .Value < 95 Then
            .Interior.ColorIndex = 36               'Highlight yellow
            .Font.ColorIndex = 12                   'Font color dark yellow
        ElseIf .Value < 90 And .Value >= 1 Then
            .Interior.ColorIndex = 46               'Highlight light red
            .Font.ColorIndex = 9                    'Font color dark red
        ElseIf .Value = 0 Then
            'alternate color if ratio=0
            If lRow Mod 2 = 0 Then
                .Interior.ColorIndex = xlNone       'Highlight no color
            Else
                .Interior.Color = RGB(230, 230, 230)    'Highlight light gray
            End If
            .Font.ColorIndex = 9                        'Font color dark red
        End If
        'Add percent sign to ratio value
        .Value = wsLineReport.Range("J" & lRow).Value & " %"
    End With
    
    'Add percentage to Pass Hi/Lo Ratios
    wsLineReport.Range("E" & lRow).Value = wsLineReport.Range("E" & lRow).Value & "  %"
    wsLineReport.Range("F" & lRow).Value = wsLineReport.Range("F" & lRow).Value & "  %"
    
    'Add percentage to Scan Hi/Lo Ratios
    wsLineReport.Range("H" & lRow).Value = wsLineReport.Range("H" & lRow).Value & "  %"
    wsLineReport.Range("I" & lRow).Value = wsLineReport.Range("I" & lRow).Value & "  %"
    
Next lRow

'Auto fit rows & columns
wsLineReport.Rows("4:4").RowHeight = 28
wsLineReport.Columns.AutoFit

ErrorExit:
On Error Resume Next
Exit Function

ErrorHandler:
Format_LineReport = False
If bCentralErrorHandler(msMODULE, sSOURCE) Then
    Stop
    Resume
Else
    Resume ErrorExit
End If
End Function

Private Function Format_LineFiveShift() As Boolean
'Error handling
Const sSOURCE As String = "Format_LineFiveShift()"
On Error GoTo ErrorHandler
Format_LineFiveShift = True

'Declare
Dim lLastRow As Long, lRow As Long, lCol As Long
Dim sPrev As String, lLineCnt As Long
lLastRow = wsLineReport.Range("B4").End(xlDown).Row

'Freeze pane
wsLineReport.Range("B5").Select
ActiveWindow.FreezePanes = True

'Right align values
With wsLineReport.Range("E5:O" & lLastRow)
    .HorizontalAlignment = xlRight
    .IndentLevel = 1
End With

'Set Start & End Period date format
wsLineReport.Range("C5:D" & lLastRow).NumberFormat = "yyyy/mm/dd hh:mm:ss"

'Alternate highlighting color
sPrev = wsLineReport.Range("B5").Value
lLineCnt = 1
For lRow = 5 To lLastRow
    
    'Alternating highlight per station
    If Not wsLineReport.Range("B" & lRow).Value = sPrev Then
        lLineCnt = lLineCnt + 1
        sPrev = wsLineReport.Range("B" & lRow).Value
    End If
    If Not lLineCnt Mod 2 = 0 Then
        wsLineReport.Range("B" & lRow & ":O" & lRow).Interior.Color = RGB(230, 230, 230)
    End If
    
    'Calculate Avg Pass Ratio & format with GA color code
    '  Green>=99, Yellow>=97 & <99, Red<97
    With wsLineReport.Range("G" & lRow)
        If .Value >= 99 Then
            .Interior.ColorIndex = 43               'Highlight green
            .Font.ColorIndex = 10                   'Font color dark green
        ElseIf .Value >= 97 And .Value < 99 Then
            .Interior.ColorIndex = 36               'Highlight yellow
            .Font.ColorIndex = 12                   'Font color dark yellow
        ElseIf .Value < 97 And .Value >= 1 Then
            .Interior.ColorIndex = 46               'Highlight light red
            .Font.ColorIndex = 9                    'Font color dark red
        ElseIf .Value = 0 Then
            'alternate color if ratio=0
            If lRow Mod 2 = 0 Then
                .Interior.ColorIndex = xlNone       'Highlight no color
            Else
                .Interior.Color = RGB(230, 230, 230)    'Highlight light gray
            End If
            .Font.ColorIndex = 9                        'Font color dark red
        End If
        'Add percent sign to ratio value
        .Value = wsLineReport.Range("G" & lRow).Value & " %"
    End With
    
    'Calculate Avg Scan Ratio & format with GA color code
    '  Green>=95, Yellow>=90 & <95, Red<90
    With wsLineReport.Range("J" & lRow)
        If .Value >= 95 Then
            .Interior.ColorIndex = 43               'Highlight green
            .Font.ColorIndex = 10                   'Font color dark green
        ElseIf .Value >= 90 And .Value < 95 Then
            .Interior.ColorIndex = 36               'Highlight yellow
            .Font.ColorIndex = 12                   'Font color dark yellow
        ElseIf .Value < 90 And .Value >= 1 Then
            .Interior.ColorIndex = 46               'Highlight light red
            .Font.ColorIndex = 9                    'Font color dark red
        ElseIf .Value = 0 Then
            'alternate color if ratio=0
            If lRow Mod 2 = 0 Then
                .Interior.ColorIndex = xlNone       'Highlight no color
            Else
                .Interior.Color = RGB(230, 230, 230)    'Highlight light gray
            End If
            .Font.ColorIndex = 9                        'Font color dark red
        End If
        'Add percent sign to ratio value
        .Value = wsLineReport.Range("J" & lRow).Value & " %"
    End With
    
    'Add percentage to Pass Hi/Lo Ratios
    wsLineReport.Range("E" & lRow).Value = wsLineReport.Range("E" & lRow).Value & "  %"
    wsLineReport.Range("F" & lRow).Value = wsLineReport.Range("F" & lRow).Value & "  %"
    
    'Add percentage to Scan Hi/Lo Ratios
    wsLineReport.Range("H" & lRow).Value = wsLineReport.Range("H" & lRow).Value & "  %"
    wsLineReport.Range("I" & lRow).Value = wsLineReport.Range("I" & lRow).Value & "  %"
Next lRow

'Auto fit rows & columns
wsLineReport.Rows("4:4").RowHeight = 28
wsLineReport.Columns.AutoFit


ErrorExit:
On Error Resume Next
Exit Function

ErrorHandler:
Format_LineFiveShift = False
If bCentralErrorHandler(msMODULE, sSOURCE) Then
    Stop
    Resume
Else
    Resume ErrorExit
End If
End Function

Private Function Template_LineReport() As Boolean
'Error handling
Const sSOURCE As String = "Template_LineReport()"
On Error GoTo ErrorHandler
Template_LineReport = True

'Declare
Dim ws As Worksheet
Dim sShiftTitle As String, sReportTitle As String, sShtName As String

'Build report sheet name
Select Case GetShift()
    Case Is = "0"
        sShtName = "All Shifts"
    Case Is = "1"
        sShtName = "1st Shift"
    Case Is = "2"
        sShtName = "2nd Shift"
    Case Is = "3"
        sShtName = "3rd Shift"
End Select
Select Case GetReport()
    Case Is = "0"
        sShtName = "Single Line Trend"
    Case Is = "1"
        sShtName = "Line - " & sShtName & " - Last Shift"
    Case Is = "2"
        sShtName = "Line - " & sShtName & " - Last5Shifts"
End Select

'Delete existing worksheet
For Each ws In Worksheets
    Application.DisplayAlerts = False
    If ws.Name = sShtName Then ws.Delete
Next ws

'Create new report worksheet
Set wsLineReport = Sheets.Add
wsLineReport.Name = sShtName
wsLineReport.Select
Application.DisplayAlerts = True

'Title1 / Plant Name
With wsLineReport.Range("B1:O1")
    .Merge
    .HorizontalAlignment = xlCenter
    .Font.Size = 12
    .Font.Bold = True
    .Value = "HMMA GA TDM"
End With
'Title 2 / Shift Description
With wsLineReport.Range("B2:O2")
    .Merge
    .HorizontalAlignment = xlCenter
    .Font.Size = 12
    .Font.Bold = True
    Select Case GetShift()
        Case Is = "0"
            sShiftTitle = "All Shifts - Line Level"
        Case Is = "1"
            sShiftTitle = "1st Shift - Line Level"
        Case Is = "2"
            sShiftTitle = "2nd Shift - Line Level"
        Case Is = "3"
            sShiftTitle = "3rd Shift - Line Level"
    End Select
    .Value = sShiftTitle
End With
'Title3 / Report Type Description
With wsLineReport.Range("B3:O3")
    .Merge
    .HorizontalAlignment = xlCenter
    .Font.Size = 12
    .Font.Bold = True
    Select Case GetReport()
        Case Is = "1"
            sReportTitle = "Last Completed Shift"
        Case Is = "2"
            sReportTitle = "Last 5 Completed Shifts"
        Case Is = "0"
            wsLineReport.Range("B2:M2").Value = "Line Level"
            sReportTitle = "Trending Period  " & GetStartDate() & " " & GetStartTime & " - " & GetEndDate & " " & GetEndTime
    End Select
    .Value = sReportTitle
End With

'Header row
With wsLineReport
    .Range("B4").Value = "Line"
    .Range("C4").Value = "Period Start"
    .Range("D4").Value = "Period End"
    .Range("E4").Value = "High Pass"
    .Range("F4").Value = "Low Pass"
    .Range("G4").Value = "Avg Pass"
    .Range("H4").Value = "High Scan"
    .Range("I4").Value = "Low Scan"
    .Range("J4").Value = "Avg Scan"
    .Range("K4").Value = "Total NG"
    .Range("L4").Value = "Torque High"
    .Range("M4").Value = "Torque Low"
    .Range("N4").Value = "Angle High"
    .Range("O4").Value = "Angle Low"
End With

'Format header row
With wsLineReport.Range("B4:O4")
    .Interior.ColorIndex = 41
    .Font.ColorIndex = 2
    .Font.Bold = True
    .Font.Size = 12
    .AutoFilter
    .Columns.AutoFit
End With

'Change header row height
wsLineReport.Rows("4:4").RowHeight = 28

ErrorExit:
On Error Resume Next
Set ws = Nothing
Exit Function

ErrorHandler:
Template_LineReport = False
If bCentralErrorHandler(msMODULE, sSOURCE) Then
    Stop
    Resume
Else
    Resume ErrorExit
End If
End Function


'Setters
'---------------
Sub SetLine(ByVal thisLine As String)
msLine = thisLine
End Sub
Sub SetStation(ByVal thisStation As String)
msStation = thisStation
End Sub
Sub SetReport(ByVal thisReport As String)
msReport = thisReport
End Sub
Sub SetShift(ByVal thisShift As String)
msShift = thisShift
End Sub
Sub SetStartDate(ByVal thisStartDate As String)
msStartDate = thisStartDate
End Sub
Sub SetStartTime(ByVal thisStartTime As String)
msStartTime = thisStartTime
End Sub
Sub SetEndDate(ByVal thisEndDate As String)
msEndDate = thisEndDate
End Sub
Sub SetEndTime(ByVal thisEndTime As String)
msEndTime = thisEndTime
End Sub

'Getters
'---------------
Function GetLine() As String
GetLine = msLine
End Function
Function GetStation() As String
GetStation = msStation
End Function
Function GetReport() As String
GetReport = msReport
End Function
Function GetShift() As String
GetShift = msShift
End Function
Function GetStartDate() As String
GetStartDate = msStartDate
End Function
Function GetStartTime() As String
GetStartTime = msStartTime
End Function
Function GetEndDate() As String
GetEndDate = msEndDate
End Function
Function GetEndTime() As String
GetEndTime = msEndTime
End Function

