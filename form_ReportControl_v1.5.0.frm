VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_ReportControl 
   Caption         =   "Report Control Panel"
   ClientHeight    =   10695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11160
   OleObjectBlob   =   "form_ReportControl_v1.5.0.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "form_ReportControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-- MODULE VARS --
'=================
Private Const msLines As String = "ALL LINES,T1-Trim Line 1,T2-Trim Line 2,T3-Trim Line 3" & _
        ",Chassis 1,Chassis Marriage,Chassis Front Sub,Chassis Sub,Chassis 2,Final 1,Final 2" & _
        ",Final 3,Final 4,Final 5,Final TF,ED Line"
Private Const msStations As String = "-:ALL: STATIONS,T1:12LA: Curtain Airbag,T1:12LB: Curtain Airbag" & _
        ",T1:13L: Curtain Airbag,T1:13R: Curtain Airbag,T2:03L: Brake Pedal,T2:04L: Trim Accel Pedal" & _
        ",T2:05L: Trans Mount Brkt,T2:07R: Eng. Earth to Body,T2:15L: Seat Belt LH" & _
        ",T2:15R: Seat Belt RH,TF:01L: -Battery Terminal,C1:10L:Fuel Tank,C2:01L:Trans Mount,C2:10L: Muffler" & _
        ",CS:04L: F Strut Knuckle L,CS:04R: F Strut Knuckle R,CS:07L: Stepbar Link L" & _
        ",CS:07R: Stepbar Link R,F1:05L: Ignition Coil Ground,F1:07L: U-Joint,F1:08L: Eng. Ground Wires" & _
        ",F1:09L: Trans Ground,F1:12R: Seat Belt Buckle,F2:07L: JBox,F2:11L: Rear Seat Anchor L" & _
        ",F4:01L: + Battery Cable,F4:02L: Seat Front Bolts Secure,F4:02R: Seat Front Bolts Secure" & _
        ",F4:04L: Strut Nuts L,F4:04R: Strut Nuts R,F4:05L: Fr. Seat Bolts L,F4:05R: Fr. Seat Bolts R" & _
        ",F4:06L: Steering Wheel,F4:08L: Wipers,F5:01R: Lower Roll Stop,F5:02L: Lower Shock ABS L" & _
        ",F5:03R: Lower Shock ABS R,ED:07R: AT Torque Conv. L,ED:08R: AT Torque Conv. R,ED:09L: Starter" & _
        ",ED:11L: Alternator B+,ED:14R: Roll Stop Brkt,ED:16L: Compressor"

Private msLineNames As String
Private msStationNames As String


'-- BUTTONS --
'==============
Private Sub btnAbout_Click()
form_About.Show
End Sub
Private Sub btnClose_Click()
Me.Hide
End Sub
Private Sub btnRunReport_Click()
On Error GoTo 0
'Show error labels
If IsNull(Reports.GetReport) Then
    Me.lblErrorMessage.Caption = "  Must select a Report Type option to continue."
    Me.lblErrorMessage.Visible = True
    Exit Sub
ElseIf IsNull(Reports.GetShift) Or Reports.GetShift = "" Then
    Me.lblErrorMessage.Caption = "  Must select a Daily Shift option to continue."
    Me.lblErrorMessage.Visible = True
    Exit Sub
Else
    Me.lblErrorMessage = False
End If

'If trending report set date/time global vars
If Me.optReportTrendingPeriod Then
    With Me
        Reports.SetStartDate (.txbStartPeriodDate.Value)
        Reports.SetStartTime (.txbStartPeriodTime.Value)
        Reports.SetEndDate (.txbEndPeriodDate.Value)
        Reports.SetEndTime (.txbEndPeriodTime.Value)
    End With
End If

'Set screen params
Application.ScreenUpdating = False
Application.Cursor = xlWait

'Call subroutine to get the report
If Me.optStationSelect.Value Then
    If Me.optReportFiveShifts.Value Then
        Call Reports.Make_StationFiveReport
    Else
        Call Reports.Make_StationReport
    End If
ElseIf Me.optLineSelect.Value Then
    If Me.optReportFiveShifts.Value Then
        Call Reports.Make_LineFiveShift
    Else
        Call Reports.Make_LineReport
    End If
End If
    
'Set screen params
Application.ScreenUpdating = True
Application.Cursor = xlDefault

If Me.chkAutoClose.Value Then
    Me.Hide
End If
Exit Sub
    
Err:
On Error Resume Next
Application.ScreenUpdating = True
Application.Cursor = xlDefault
MsgBox Err.Number & " : " & Err.Description, vbCritical + vbOKOnly, "Error"

End Sub

'-- REPORT OPTIONS --
'  Option parameters 0=Trending, 1=LastShift, 2=FiveShifts
'=====================
Private Sub optReportLastShift_Change()
If Me.optReportLastShift.Value Then
    Reports.SetReport ("1")
    
    'Jump over weekend dates
    If Format(Now() - 1, "ddd") = "Sun" Then
        Reports.SetStartDate (Format(Now() - 3, "yyyy/mm/dd"))
        Reports.SetEndDate (Format(Now() - 3, "yyyy/mm/dd"))
    ElseIf Format(Now() - 1, "ddd") = "Sat" Then
        Reports.SetStartDate (Format(Now() - 2, "yyyy/mm/dd"))
        Reports.SetEndDate (Format(Now() - 2, "yyyy/mm/dd"))
    Else
        Reports.SetStartDate (Format(Now() - 1, "yyyy/mm/dd"))
        Reports.SetEndDate (Format(Now() - 1, "yyyy/mm/dd"))
    End If
    
    'Adjust end date dependent on select shift
    If Reports.GetShift() = 0 Or Reports.GetShift() = 3 Then
        Reports.SetEndDate (Format(Format(Reports.GetStartDate, "yyyy/mm") & "/" & Format(Reports.GetStartDate, "dd") + 1, "yyyy/mm/dd"))
    End If
End If
'Enable Daily shifts
Me.frmDailyShift.Enabled = True
End Sub
Private Sub optReportFiveShifts_Change()
If Me.optReportFiveShifts.Value Then
    Reports.SetReport ("2")
    Reports.SetStartDate (Format(Now() - 7, "yyyy/mm/dd"))
    Reports.SetEndDate (Format(Now() - 1, "yyyy/mm/dd"))
End If
'Enable Daily shifts
Me.frmDailyShift.Enabled = True
End Sub
Private Sub optReportTrendingPeriod_Change()
If Me.optReportTrendingPeriod.Value Then
    Reports.SetReport ("0")
    'Configure trending period textbox defaults
    With Me.txbStartPeriodDate
        .Value = Reports.GetStartDate()
        .Enabled = True
        .SetFocus
    End With
    With Me.txbStartPeriodTime
        .Value = Reports.GetStartTime()
        .Enabled = True
    End With
    With Me.txbEndPeriodDate
        .Value = Reports.GetEndDate()
        .Enabled = True
    End With
    With Me.txbEndPeriodTime
        .Value = Reports.GetEndTime()
        .Enabled = True
    End With
    'Disable Daily shifts
    Me.frmDailyShift.Enabled = False
Else
    'Set default date/time in textbox
    Me.txbStartPeriodDate.Enabled = False
    Me.txbStartPeriodTime.Enabled = False
    Me.txbEndPeriodDate.Enabled = False
    Me.txbEndPeriodTime.Enabled = False
        
    'Hide error labels
    Me.lblStartPeriodError.Visible = False
    Me.lblEndPeriodError.Visible = False

End If
End Sub


'-- SHIFT OPTIONS --
'   Shift parameters 0=All, 1=First, 2=Second, 3=Third
'=====================
Private Sub optShiftAll_Change()
If Me.optShiftAll.Value Then
    Reports.SetShift ("0")
    Reports.SetStartTime ("06:45:00")
    Reports.SetEndTime ("06:45:00")
    Reports.SetEndDate (Format(Format(Reports.GetStartDate, "yyyy/mm") & "/" & Format(Reports.GetStartDate, "dd") + 1, "yyyy/mm/dd"))
End If
End Sub
Private Sub optShiftFirst_Change()
If Me.optShiftFirst.Value Then
    Reports.SetShift ("1")
    Reports.SetStartTime ("06:45:00")
    Reports.SetEndTime ("14:45:00")
    Reports.SetEndDate (Reports.GetStartDate())
End If
End Sub
Private Sub optShiftSecond_Change()
If Me.optShiftSecond.Value Then
    Reports.SetShift ("2")
    Reports.SetStartTime ("14:45:00")
    Reports.SetEndTime ("22:45:00")
    Reports.SetEndDate (Reports.GetStartDate())
End If
End Sub
Private Sub optShiftThird_Change()
Dim s As String
If Me.optShiftThird.Value Then
    Reports.SetShift ("3")
    Reports.SetStartTime ("22:45:00")
    Reports.SetEndTime ("06:45:00")
    Reports.SetEndDate (Format(Format(Reports.GetStartDate, "yyyy/mm") & "/" & Format(Reports.GetStartDate, "dd") + 1, "yyyy/mm/dd"))
End If
End Sub


'-- LINE SELECT --
'=================
Private Sub optLineSelect_Change()
Dim iIndex As Integer
If Me.optLineSelect.Value Then
    'Enable list
    Me.lsbLines.Enabled = True
Else
    'Disable list
    Me.lsbLines.Enabled = False
    
    'Unselect items
    For iIndex = 0 To Me.lsbLines.ListCount - 1
        If Me.lsbLines.Selected(iIndex) Then Me.lsbLines.Selected(iIndex) = False
    Next iIndex
End If
End Sub
Private Sub lsbLines_Change()
Dim iIndex As Integer
Dim sLines As String
sLines = ""

For iIndex = 0 To Me.lsbLines.ListCount - 1
    If Me.lsbLines.Selected(iIndex) Then
        If sLines = "" Then
            sLines = Me.lsbLines.List(iIndex)
        Else
            sLines = sLines & "," & Me.lsbLines.List(iIndex)
        End If
    End If
Next iIndex

If Mid(sLines, 1, 3) = "ALL" Then
    For iIndex = 0 To Me.lsbLines.ListCount - 1
        If Mid(Me.lsbLines.List(iIndex), 1, 3) = "ALL" Then
            Me.lsbLines.Selected(iIndex) = True
        Else
            If Me.lsbLines.Selected(iIndex) Then Me.lsbLines.Selected(iIndex) = False
        End If
    Next iIndex
    
    sLines = "All"
End If

msLineNames = sLines
Reports.SetLine (msLineNames)
End Sub


'-- STATION SELECT --
'=====================
Private Sub optStationSelect_Change()
Dim iIndex As Integer
If Me.optStationSelect.Value Then
    Me.lsbStations.Enabled = True
Else
    'Disable list
    Me.lsbStations.Enabled = False
End If
End Sub
Private Sub lsbStations_Change()
Dim iIndex As Integer
Dim sStations As String
sStations = ""

For iIndex = 0 To Me.lsbStations.ListCount - 1
    If Me.lsbStations.Selected(iIndex) Then
        If sStations = "" Then
            sStations = Me.lsbStations.List(iIndex)
        Else
            sStations = sStations & "," & Me.lsbStations.List(iIndex)
        End If
    End If
Next iIndex

If Mid(sStations, 1, 5) = "--ALL" Then
    For iIndex = 0 To Me.lsbStations.ListCount - 1
        If Mid(Me.lsbStations.List(iIndex), 1, 5) = "--ALL" Then
            Me.lsbStations.Selected(iIndex) = True
        Else
            If Me.lsbStations.Selected(iIndex) Then Me.lsbStations.Selected(iIndex) = False
        End If
    Next iIndex
    
    sStations = "All"
End If

Reports.SetStation (sStations)

End Sub


'-- TRENDING PERIOD --
'=====================
Private Sub txbStartPeriodDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)
Dim sStartDate As String
sStartDate = Trim(Me.txbStartPeriodDate.Value)

'Show/hide error label
If VerifyDate(sStartDate) Then
    Me.lblStartPeriodError.Visible = False
    Reports.SetStartDate (sStartDate)
Else
    Me.lblStartPeriodError.Visible = True
End If

End Sub
Private Sub txbEndPeriodDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)
Dim sEndDate As String
sEndDate = Trim(Me.txbEndPeriodDate.Value)

'Show/hide error label
If VerifyDate(sEndDate) Then
    Me.lblEndPeriodError.Visible = False
    Reports.SetEndDate (sEndDate)
Else
    Me.lblEndPeriodError.Visible = True
End If

End Sub
Private Sub txbStartPeriodTime_Exit(ByVal Cancel As MSForms.ReturnBoolean)
Dim sStartTime As String
sStartTime = Trim(Me.txbStartPeriodTime.Value)

'Show/hide error label
If VerifyTime(sStartTime) Then
    Me.lblStartPeriodError.Visible = False
    Reports.SetStartTime (sStartTime)
Else
    Me.lblStartPeriodError.Visible = True
End If

End Sub
Private Sub txbEndPeriodTime_Exit(ByVal Cancel As MSForms.ReturnBoolean)
Dim sEndTime As String
sEndTime = Trim(Me.txbEndPeriodTime.Value)

'Show/hide error label
If VerifyTime(sEndTime) Then
    Me.lblEndPeriodError.Visible = False
    Reports.SetEndTime (sEndTime)
Else
    Me.lblEndPeriodError.Visible = True
End If

End Sub

'-- INTERN FUNCTIONS --
'========================
Private Function VerifyDate(ByVal thisDate As String) As Boolean
Dim bValid As Boolean
bValid = True

'Do nothing if field is blank
If thisDate = "" Then GoTo QuickExit

'Show error label if count not equal 10
If Len(thisDate) = 10 Then
    'Error if "/" not in correct places
    If Not (Mid(thisDate, 5, 1) = "/") Or _
            Not (Mid(thisDate, 8, 1) = "/") Or _
            Not (IsNumeric(Mid(thisDate, 1, 4))) Or _
            Not (IsNumeric(Mid(thisDate, 6, 2))) Or _
            Not (IsNumeric(Mid(thisDate, 9, 2))) Then
        bValid = False
    End If
Else
    bValid = False
End If

QuickExit:
VerifyDate = bValid

End Function
Private Function VerifyTime(ByVal thisTime As String) As Boolean
Dim bValid As Boolean
bValid = True

'Do nothing if field is blank
If thisTime = "" Then GoTo QuickExit

'Show error label if count is not equal to 8
If Len(thisTime) = 8 Then
    'Error if ":" not in correct places, or if not numeric
    If Not (Mid(thisTime, 3, 1) = ":") Or _
            Not (Mid(thisTime, 6, 1) = ":") Or _
            Not (IsNumeric(Left(thisTime, 2))) Or _
            Not (IsNumeric(Mid(thisTime, 4, 2))) Or _
            Not (IsNumeric(Mid(thisTime, 7, 2))) Then
        bValid = False
    End If
Else
    bValid = False
End If

QuickExit:
VerifyTime = bValid

End Function


'-- INITIALIZE --
'================
Private Sub UserForm_Initialize()
Dim arLines() As String
Dim arStations() As String, arSta() As String
Dim sStationName As String
Dim iLineIndex As Integer, iStaIndex As Integer

'Add line list items
arLines() = Split(msLines, ",")
For iLineIndex = LBound(arLines) To UBound(arLines)
    Me.lsbLines.AddItem arLines(iLineIndex)
Next iLineIndex

'Add stations to list items
arStations() = Split(msStations, ",")
For iStaIndex = LBound(arStations) To UBound(arStations)
    arSta() = Split(arStations(iStaIndex), ":")
    sStationName = arSta(0) & "-" & arSta(1) & arSta(2)
    Me.lsbStations.AddItem sStationName
Next iStaIndex

'Enable station listbox
Me.lsbStations.Enabled = True
Me.lsbStations.Selected(0) = True

'Disable line listbox
Me.lsbLines.Enabled = False
Me.lsbLines.Selected(0) = True

'Start & End Date defaults
'/ Jump over weekend dates
If Format(Now() - 1, "ddd") = "Sun" Then
    Reports.SetStartDate (Format(Now() - 3, "yyyy/mm/dd"))
    Reports.SetEndDate (Format(Now() - 2, "yyyy/mm/dd"))
ElseIf Format(Now() - 1, "ddd") = "Sat" Then
    Reports.SetStartDate (Format(Now() - 2, "yyyy/mm/dd"))
    Reports.SetEndDate (Format(Now() - 1, "yyyy/mm/dd"))
Else
    Reports.SetStartDate (Format(Now() - 1, "yyyy/mm/dd"))
    Reports.SetEndDate (Format(Now(), "yyyy/mm/dd"))
End If

Reports.SetStartTime ("06:45:00")
Reports.SetEndTime ("06:45:00")
Reports.SetReport ("1")            'Report default
Reports.SetShift ("0")             'Shift default
Reports.SetStation ("All")         'Station default
Reports.SetLine ("All")            'Line default

End Sub
