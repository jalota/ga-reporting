Attribute VB_Name = "Dashboard"
Option Explicit

Private Const msMODULE As String = "Dashboard"

Private wsDashboard As Worksheet
Private rsDash As ADODB.Recordset
Private cnSqlServer As SQLServerConn
Private objChart As ChartObject


'Dashboard
'--------------
Sub Make_Dashboard()
'Error handling
Const sSOURCE As String = "Make_Dashboard()"
Const bDebug As Boolean = False
On Error GoTo ErrorHandler
Application.ScreenUpdating = False

'Allow user to halt program
Application.EnableCancelKey = xlErrorHandler

'Declare
Dim ws As Worksheet

'Initialize sql server connection
Set cnSqlServer = New SQLServerConn
cnSqlServer.ConnOpen

'Check if Dashboard worksheet exists
For Each ws In Worksheets
    Application.DisplayAlerts = False
    If ws.Name = "Dashboard" Then ws.Delete
Next ws

'Create Dashboard worksheet if not exists
Set wsDashboard = Worksheets.Add
wsDashboard.Name = "Dashboard"
wsDashboard.Select
Application.DisplayAlerts = True

Start:
'Change background white
wsDashboard.Columns.Interior.ColorIndex = 2

'Set title
With Range("E2:N2")
    .Merge
    .HorizontalAlignment = xlCenter
    .Value = "Production Performance Dashboard"
    .Font.Bold = True
    .Font.Size = 18
End With
With Range("E3:N3")
    .Merge
    .HorizontalAlignment = xlCenter
    .Value = "( " & Format(Now() - 1, "yyyy/mm/dd") & " )"
    .Font.Bold = True
    .Font.Size = 14
End With

'Add charts
If Not Chart_Monthly() Then Err.Raise glHANDLED_ERROR
If Not Chart_Daily() Then Err.Raise glHANDLED_ERROR
If Not Chart_Line() Then Err.Raise glHANDLED_ERROR

ErrorExit:
Application.ScreenUpdating = True
cnSqlServer.ConnClose
Set ws = Nothing
Set wsDashboard = Nothing
Exit Sub

ErrorHandler:
If bCentralErrorHandler(msMODULE, sSOURCE, , True) Then
    Stop
    Resume
Else
    Resume ErrorExit
End If
End Sub


Private Function Chart_Monthly() As Boolean
'Error handling
Const sSOURCE As String = "Chart_Monthly()"
On Error GoTo ErrorHandler
Chart_Monthly = True

'Declare
Dim lRow As Long

'Retrieve Facility Monthly data
Set rsDash = New ADODB.Recordset
rsDash.Open Sql.FacilityMonthly, cnSqlServer.Conn, adOpenStatic

If Not rsDash.BOF And Not rsDash.EOF Then
    With wsDashboard
        .Range("B8").Value = "Pass"
        .Range("C8").Value = "Scan"
        .Range("A" & (rsDash.RecordCount + 9)).Value = "Total"
        .Range("A9").CopyFromRecordset rsDash
        .Range("B" & (rsDash.RecordCount + 9)).Value = Application.WorksheetFunction.Average(Range("B9:B" & (rsDash.RecordCount + 9))) & "%"
        .Range("C" & (rsDash.RecordCount + 9)).Value = Application.WorksheetFunction.Average(Range("C9:C" & (rsDash.RecordCount + 9))) & "%"
        
        For lRow = 0 To rsDash.RecordCount - 1
            .Range("B" & (lRow + 9)).Value = .Range("B" & (lRow + 9)).Value & "%"
            .Range("C" & (lRow + 9)).Value = .Range("C" & (lRow + 9)).Value & "%"
        Next lRow
        
        .Range("A8:C" & (rsDash.RecordCount + 9)).Font.ColorIndex = 2
    End With
End If
'Add & define Facilty Monthly chart
Set objChart = ActiveSheet.ChartObjects.Add _
    (Left:=10, Width:=575, Top:=85, Height:=200)
With objChart
    .Chart.SetSourceData Source:=Sheets("Dashboard").Range("A8:C" & (rsDash.RecordCount + 9))
    .Chart.ChartType = xlColumnClustered
    .Chart.ChartStyle = 10
    .Chart.SetElement (msoElementPrimaryValueAxisNone)
    .Chart.SetElement (msoElementPrimaryValueGridLinesNone)
    .Chart.SetElement (msoElementDataLabelOutSideEnd)
    .Chart.SetElement (msoElementChartTitleAboveChart)
    .Chart.HasTitle = True
    .Chart.ChartTitle.Text = "Total Status Board (Monthly)"
    .RoundedCorners = True
    .Chart.ChartArea.Border.LineStyle = xlNone
End With

ErrorExit:
On Error Resume Next
rsDash.Close
Set rsDash = Nothing
Set objChart = Nothing
Exit Function

ErrorHandler:
Chart_Monthly = False
If bCentralErrorHandler(msMODULE, sSOURCE) Then
    Stop
    Resume
Else
    Resume ErrorExit
End If
End Function

Private Function Chart_Daily() As Boolean
'Error handling
Const sSOURCE As String = "Chart_Daily()"
On Error GoTo ErrorHandler
Chart_Daily = True

'Declare
Dim i As Integer
Dim sStart As String, sEnd As String, sShift As String

'Calculate date range for previous 9 of 10 completed shifts
If Format(Now(), "ddd") = "Mon" Then
    sEnd = Format(Now() - 3, "yyyy-mm-dd")
ElseIf Format(Now(), "ddd") = "Sun" Then
    sEnd = Format(Now() - 2, "yyyy-mm-dd")
Else
    sEnd = Format(Now() - 1, "yyyy-mm-dd")
End If
For i = 10 To 19
    If Application.WorksheetFunction.NetworkDays(Format(Now() - i, "yyyy-mm-dd"), sEnd) = 9 Then
        sStart = Format(Now() - i, "yyyy-mm-dd")
    End If
Next i

'Day 1~9
Set rsDash = New ADODB.Recordset
rsDash.Open Sql.DailyStatus(sStart, sEnd, "%"), cnSqlServer.Conn, adOpenStatic
If Not rsDash.BOF And Not rsDash.EOF Then
    With wsDashboard
        .Range("B24").Value = "Pass"
        .Range("C24").Value = "Scan"
        .Range("A25").Value = "Day 1~9"
        .Range("B25").Value = rsDash.Fields("PassRatio").Value & "%"
        .Range("C25").Value = rsDash.Fields("ScanRatio").Value & "%"
    End With
End If
If Not rsDash Is Nothing Then rsDash.Close

'Day 10
If Format(Now(), "ddd") = "Mon" Then
    sStart = Format(Now() - 3, "yyyy-mm-dd")
    sEnd = Format(Now() - 3, "yyyy-mm-dd")
ElseIf Format(Now(), "ddd") = "Sun" Then
    sStart = Format(Now() - 2, "yyyy-mm-dd")
    sEnd = Format(Now() - 2, "yyyy-mm-dd")
Else
    sStart = Format(Now() - 1, "yyyy-mm-dd")
    sEnd = Format(Now() - 1, "yyyy-mm-dd")
End If
rsDash.Open Sql.DailyStatus(sStart, sEnd, "%"), cnSqlServer.Conn, adOpenStatic
If Not rsDash.BOF And Not rsDash.EOF Then
    With wsDashboard
        .Range("A26").Value = "Day 10"
        .Range("B26").Value = rsDash.Fields("PassRatio").Value & "%"
        .Range("C26").Value = rsDash.Fields("ScanRatio").Value & "%"
    End With
End If
If Not rsDash Is Nothing Then rsDash.Close

'Shift A - Last completed shift
rsDash.Open Sql.DailyStatus(sStart, sEnd, "A"), cnSqlServer.Conn, adOpenStatic
If Not rsDash.BOF And Not rsDash.EOF Then
    With wsDashboard
        .Range("A27").Value = "Shift A"
        .Range("B27").Value = rsDash.Fields("PassRatio").Value & "%"
        .Range("C27").Value = rsDash.Fields("ScanRatio").Value & "%"
    End With
End If
If Not rsDash Is Nothing Then rsDash.Close

'Shift B - Last completed shift
rsDash.Open Sql.DailyStatus(sStart, sEnd, "B"), cnSqlServer.Conn, adOpenStatic
If Not rsDash.BOF And Not rsDash.EOF Then
    With wsDashboard
        .Range("A28").Value = "Shift B"
        .Range("B28").Value = rsDash.Fields("PassRatio").Value & "%"
        .Range("C28").Value = rsDash.Fields("ScanRatio").Value & "%"
    End With
End If
If Not rsDash Is Nothing Then rsDash.Close

'Shift C - Last completed shift
rsDash.Open Sql.DailyStatus(sStart, sEnd, "C"), cnSqlServer.Conn, adOpenStatic
If Not rsDash.BOF And Not rsDash.EOF Then
    With wsDashboard
        .Range("A29").Value = "Shift C"
        .Range("B29").Value = rsDash.Fields("PassRatio").Value & "%"
        .Range("C29").Value = rsDash.Fields("ScanRatio").Value & "%"
    End With
End If
If Not rsDash Is Nothing Then rsDash.Close

'Configure chart object
Set objChart = ActiveSheet.ChartObjects.Add _
    (Left:=10, Width:=575, Top:=300, Height:=200)
With objChart
    .Chart.SetSourceData Source:=Sheets("Dashboard").Range("A24:C29")
    .Chart.ChartType = xlColumnClustered
    .Chart.ChartStyle = 10
    .Chart.SetElement (msoElementPrimaryValueAxisNone)
    .Chart.SetElement (msoElementPrimaryValueGridLinesNone)
    .Chart.SetElement (msoElementDataLabelOutSideEnd)
    .Chart.SetElement (msoElementChartTitleAboveChart)
    .Chart.HasTitle = True
    .Chart.ChartTitle.Text = "Total Status Board (Daily)"
    .RoundedCorners = True
    .Chart.ChartArea.Border.LineStyle = xlNone
End With

'Change chart data font to white
wsDashboard.Range("A24:C29").Font.ColorIndex = 2

ErrorExit:
On Error Resume Next
rsDash.Close
Set rsDash = Nothing
Set objChart = Nothing
Exit Function

ErrorHandler:
Chart_Daily = False
If bCentralErrorHandler(msMODULE, sSOURCE) Then
    Stop
    Resume
Else
    Resume ErrorExit
End If
End Function

Private Function Chart_Line() As Boolean
'Error handling
Const sSOURCE As String = "Chart_Line()"
On Error GoTo ErrorHandler
Chart_Line = True

'Declare
Dim lRow As Long

Set rsDash = New ADODB.Recordset
rsDash.Open Sql.LineStatus, cnSqlServer.Conn, adOpenStatic
If Not rsDash.BOF And Not rsDash.EOF Then
    With wsDashboard
        'Set data header fields
        .Range("B39").Value = "Pass"
        .Range("C39").Value = "Scan"
        
        'Copy data to worksheet
        .Range("A40").CopyFromRecordset rsDash
        
        'Add `%` to values
        For lRow = 0 To rsDash.RecordCount - 1
            .Range("B" & (lRow + 40)).Value = .Range("B" & (lRow + 40)).Value & "%"
            .Range("C" & (lRow + 40)).Value = .Range("C" & (lRow + 40)).Value & "%"
        Next lRow
        
        'Change data font color to white
        .Range("A39:C" & (rsDash.RecordCount + 39)).Font.ColorIndex = 2
    End With
End If

'Add & define Line Status chart
Set objChart = ActiveSheet.ChartObjects.Add _
    (Left:=10, Width:=575, Top:=515, Height:=200)
With objChart
    .Chart.SetSourceData Source:=Sheets("Dashboard").Range("A39:C" & (rsDash.RecordCount + 39))
    .Chart.ChartType = xlColumnClustered
    .Chart.ChartStyle = 10
    .Chart.SetElement (msoElementPrimaryValueAxisNone)
    .Chart.SetElement (msoElementPrimaryValueGridLinesNone)
    .Chart.SetElement (msoElementDataLabelOutSideEnd)
    .Chart.SetElement (msoElementChartTitleAboveChart)
    .Chart.HasTitle = True
    .Chart.ChartTitle.Text = "Line Status Board"
    .RoundedCorners = True
    .Chart.ChartArea.Border.LineStyle = xlNone
End With

ErrorExit:
On Error Resume Next
rsDash.Close
Set rsDash = Nothing
Set objChart = Nothing
Exit Function

ErrorHandler:
Chart_Line = False
If bCentralErrorHandler(msMODULE, sSOURCE) Then
    Stop
    Resume
Else
    Resume ErrorExit
End If
End Function
