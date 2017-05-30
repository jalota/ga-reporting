Attribute VB_Name = "ToolCharts"
Option Explicit

'Private Vars
'--------------
Private Const msMODULE As String = "ToolCharts"

Private wsToolCharts As Worksheet
Private rsChart As ADODB.Recordset
Private cnSqlServer As SQLServerConn
Private objChart As ChartObject

Private msStartChart1 As String
Private msEndChart1 As String
Private msStartChart2 As String
Private msEndChart2 As String
Private msStationChart1 As String
Private msStationChart2 As String
Private msShowChart1 As Boolean
Private msShowChart2 As Boolean


Sub Make_ToolChart()
'Error handling
Const sSOURCE As String = "Make_ToolChart()"
Const bDebug As Boolean = False
On Error GoTo ErrorHandler
Application.ScreenUpdating = False

'Allow user to halt program
Application.EnableCancelKey = xlErrorHandler

'Declare
Dim ws As Worksheet

'Check if ToolCharts worksheet exists
For Each ws In Worksheets
    Application.DisplayAlerts = False
    If ws.Name = "Tool Charts" Then ws.Delete
Next ws

'Create ToolCharts worksheet
Set wsToolCharts = Worksheets.Add
wsToolCharts.Name = "Tool Charts"
wsToolCharts.Select
Application.DisplayAlerts = True

'Change background white
wsToolCharts.Columns.Interior.ColorIndex = 2

'Set title
With Range("A2:N2")
    .Merge
    .HorizontalAlignment = xlCenter
    .Value = "Select Tool Performance Charting"
    .Font.Bold = True
    .Font.Size = 18
End With

'Connect to sql server
Set cnSqlServer = New SQLServerConn
cnSqlServer.ConnOpen

'Add selected charts
If GetShowChart1 Then
    If Not Chart_SingleTool() Then Err.Raise glHANDLED_ERROR
End If
If GetShowChart2 Then
    If Not Chart_MultiTool() Then Err.Raise glHANDLED_ERROR
End If

ErrorExit:
On Error Resume Next
Application.ScreenUpdating = True
cnSqlServer.ConnClose
Set ws = Nothing
Exit Sub

ErrorHandler:
If bCentralErrorHandler(msMODULE, sSOURCE, , True) Then
    Stop
    Resume
Else
    Resume ErrorExit
End If
End Sub

Private Function Chart_SingleTool() As Boolean
'Error handling
Const sSOURCE As String = "Chart_SingleTool()"
On Error GoTo ErrorHandler
Chart_SingleTool = True

'Declare
Dim lRow As Long

'Retrieve data
Set rsChart = New ADODB.Recordset
rsChart.Open Sql.SingleToolChart(GetStartChart1, GetEndChart1, GetStationChart1), cnSqlServer.Conn, adOpenStatic
If Not rsChart.BOF And Not rsChart.EOF Then
    With wsToolCharts
        'Add data headers to worksheet
        .Range("A6").Value = "Shift"
        .Range("B6").Value = "Pass"
        .Range("C6").Value = "Scan"
        
        'Add retreived data to worksheet
        .Range("A7").CopyFromRecordset rsChart
        
        'Add `%` to values
        For lRow = 0 To rsChart.RecordCount - 1
            .Range("B" & (7 + lRow)).Value = .Range("B" & (7 + lRow)).Value & "%"
            .Range("C" & (7 + lRow)).Value = .Range("C" & (7 + lRow)).Value & "%"
        Next lRow
        
        'Change font color to white
        .Range("A6:C" & (rsChart.RecordCount + 6)).Font.ColorIndex = 2
    End With
End If

'Add & define Single Tool chart
Set objChart = ActiveSheet.ChartObjects.Add _
    (Left:=10, Width:=600, Top:=85, Height:=225)
With objChart
    .Chart.ChartType = xlLineMarkers
    .Chart.SetSourceData Source:=Sheets("Tool Charts").Range("A6:C" & (rsChart.RecordCount + 6))
    .Chart.ChartStyle = 10
    .Chart.SetElement (msoElementDataLabelTop)
    .Chart.SetElement (msoElementChartTitleAboveChart)
    .Chart.HasTitle = True
    .Chart.ChartTitle.Text = GetStationChart1
    .RoundedCorners = True
    .Chart.ChartArea.Border.LineStyle = xlNone
End With

ErrorExit:
On Error Resume Next
rsChart.Close
Set rsChart = Nothing
Set objChart = Nothing
Exit Function

ErrorHandler:
Chart_SingleTool = False
If bCentralErrorHandler(msMODULE, sSOURCE) Then
    Stop
    Resume
Else
    Resume ErrorExit
End If
End Function

Private Function Chart_MultiTool() As Boolean
'Error handling
Const sSOURCE As String = "Chart_MultiTool()"
Const bDebug As Boolean = False
On Error GoTo ErrorHandler
Chart_MultiTool = True

'Declare
Dim lRow As Long, iTop As Integer
Dim sSql As String, sStation As String
Dim arSta() As String

'Compile sql
sSql = ""
arSta() = Split(GetStationChart2, ",")
For lRow = LBound(arSta) To UBound(arSta)
    sStation = sStation & " '" & arSta(lRow) & "',"
Next lRow
sStation = Left(sStation, Len(sStation) - 1)
sSql = Sql.MultiToolChart(GetStartChart2, GetEndChart2, sStation)

'/* DEBUG */
If bDebug Then
    Debug.Print sSql
    GoTo ErrorExit
End If

'Retrieve data
Set rsChart = New ADODB.Recordset
rsChart.Open sSql, cnSqlServer.Conn, adOpenStatic
If Not rsChart.BOF And Not rsChart.EOF Then
    With wsToolCharts
        'Add data headers to worksheet
        '.Range("E6").Value = "Station"
        .Range("F6").Value = "Pass"
        .Range("G6").Value = "Scan"
        
        'Add retreived data to worksheet
        .Range("E7").CopyFromRecordset rsChart
        
        'Add `%` to values
        For lRow = 0 To rsChart.RecordCount - 1
            .Range("F" & (7 + lRow)).Value = .Range("F" & (7 + lRow)).Value & "%"
            .Range("G" & (7 + lRow)).Value = .Range("G" & (7 + lRow)).Value & "%"
        Next lRow
        
        'Change font color to white
        .Range("E6:G" & (rsChart.RecordCount + 7)).Font.ColorIndex = 2
    End With
End If

'Set chart location depending if SingleTool chart shows
If GetShowChart1 Then
    iTop = 350
Else
    iTop = 85
End If

'Add & define Single Tool chart
Set objChart = ActiveSheet.ChartObjects.Add _
    (Left:=10, Width:=575, Top:=iTop, Height:=200)
With objChart
    .Chart.ChartType = xlLineMarkers
    .Chart.SetSourceData Source:=Sheets("Tool Charts").Range("E6:G" & (rsChart.RecordCount + 6))
    .Chart.ChartStyle = 10
    .Chart.SetElement (msoElementDataLabelTop)
    .Chart.SetElement (msoElementChartTitleAboveChart)
    .Chart.HasTitle = True
    .Chart.ChartTitle.Text = "(" & GetStartChart2 & ") to (" & GetEndChart2 & ")"
    .RoundedCorners = True
    .Chart.ChartArea.Border.LineStyle = xlNone
End With

ErrorExit:
On Error Resume Next
rsChart.Close
Set rsChart = Nothing
Set objChart = Nothing
Exit Function

ErrorHandler:
Chart_MultiTool = False
If bCentralErrorHandler(msMODULE, sSOURCE) Then
    Stop
    Resume
Else
    Resume ErrorExit
End If
End Function

'Setters
'---------------
Sub SetStartChart1(ByVal thisDate As String)
msStartChart1 = thisDate
End Sub
Sub SetEndChart1(ByVal thisDate As String)
msEndChart1 = thisDate
End Sub
Sub SetStartChart2(ByVal thisDate As String)
msStartChart2 = thisDate
End Sub
Sub SetEndChart2(ByVal thisDate As String)
msEndChart2 = thisDate
End Sub
Sub SetStationChart1(ByVal thisStation As String)
msStationChart1 = thisStation
End Sub
Sub SetStationChart2(ByVal thisStation As String)
msStationChart2 = thisStation
End Sub
Sub SetShowChart1(ByVal bShow As Boolean)
msShowChart1 = bShow
End Sub
Sub SetShowChart2(ByVal bShow As Boolean)
msShowChart2 = bShow
End Sub

'Getters
'---------------
Function GetStartChart1() As String
GetStartChart1 = msStartChart1
End Function
Function GetEndChart1() As String
GetEndChart1 = msEndChart1
End Function
Function GetStartChart2() As String
GetStartChart2 = msStartChart2
End Function
Function GetEndChart2() As String
GetEndChart2 = msEndChart2
End Function
Function GetStationChart1() As String
GetStationChart1 = msStationChart1
End Function
Function GetStationChart2() As String
GetStationChart2 = msStationChart2
End Function
Function GetShowChart1() As Boolean
GetShowChart1 = msShowChart1
End Function
Function GetShowChart2() As Boolean
GetShowChart2 = msShowChart2
End Function


