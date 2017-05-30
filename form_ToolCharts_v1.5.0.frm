VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_ToolCharts 
   Caption         =   "UserForm1"
   ClientHeight    =   9705
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12750
   OleObjectBlob   =   "form_ToolCharts_v1.5.0.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_ToolCharts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const msMODULE As String = "form_ToolCharts"
Private Const msStations As String = "T1:12LA: Curtain Airbag,T1:12LB: Curtain Airbag" & _
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

Private marStations() As String


' -- BUTTONS --
'===============
Private Sub btnAbout_Click()
form_About.Show
End Sub
Private Sub btnClose_Click()
Me.Hide
End Sub
Private Sub btnRun_Click()
Dim sStations As String
Dim lIndex As Long

'Set start & end dates
ToolCharts.SetStartChart1 (Me.txtStartDateChart1.Value)
ToolCharts.SetEndChart1 (Me.txtEndDateChart1.Value)
ToolCharts.SetStartChart2 (Me.txtStartDateChart2.Value)
ToolCharts.SetEndChart2 (Me.txtEndDateChart2.Value)

'Set show charts
ToolCharts.SetShowChart1 (Me.chbShowChart1.Value)
ToolCharts.SetShowChart2 (Me.chbShowChart2.Value)

'Set chart1 station selection
ToolCharts.SetStationChart1 (Me.lsbStationChart1.Value)

'Set chart2 station selection
For lIndex = 0 To Me.lsbStationChart2.ListCount - 1
    If Me.lsbStationChart2.Selected(lIndex) Then
        If sStations = "" Then
            sStations = Me.lsbStationChart2.List(lIndex)
        Else
            sStations = sStations & "," & Me.lsbStationChart2.List(lIndex)
        End If
    End If
Next lIndex
ToolCharts.SetStationChart2 (sStations)

'Call subroutine to get charts
Call ToolCharts.Make_ToolChart
Me.Hide

End Sub

' -- ENABLE/DISABLE CHARTS OBJECTS --
'=====================================
Private Sub chbShowChart1_Click()
With Me
    If .chbShowChart1.Value Then
        'Enable chart1 objects
        .txtStartDateChart1.Enabled = True
        .txtEndDateChart1.Enabled = True
        .lsbStationChart1.Enabled = True
    Else
        'Disable chart1 objects
        .txtStartDateChart1.Enabled = False
        .txtEndDateChart1.Enabled = False
        .lsbStationChart1.Enabled = False
    End If
End With
End Sub
Private Sub chbShowChart2_Click()
With Me
    If .chbShowChart2.Value Then
        'Enable chart2 objects
        .txtStartDateChart2.Enabled = True
        .txtEndDateChart2.Enabled = True
        .lsbStationChart2.Enabled = True
    Else
        'Disable chart2 objects
        .txtStartDateChart2.Enabled = False
        .txtEndDateChart2.Enabled = False
        .lsbStationChart2.Enabled = False
    End If
End With
End Sub

' -- PERIOD DATE VERIFICATION --
'================================
Private Sub txtStartDateChart1_AfterUpdate()
Dim sDate As String
sDate = Trim(Me.txtStartDateChart1.Value)
'Show/hide error label
If VerifyDate(sDate) Then
    Me.lblErrStartChart1.Visible = False
    Me.btnRun.Enabled = True
Else
    Me.lblErrStartChart1.Visible = True
    Me.btnRun.Enabled = False
End If
End Sub
Private Sub txtEndDateChart1_AfterUpdate()
Dim sDate As String
sDate = Trim(Me.txtEndDateChart1.Value)
'Show/hide error label
If VerifyDate(sDate) Then
    Me.lblErrEndChart1.Visible = False
    Me.btnRun.Enabled = True
Else
    Me.lblErrEndChart1.Visible = True
    Me.btnRun.Enabled = False
End If
End Sub
Private Sub txtStartDateChart2_AfterUpdate()
Dim sDate As String
sDate = Trim(Me.txtStartDateChart2.Value)
'Show/hide error label
If VerifyDate(sDate) Then
    Me.lblErrStartChart2.Visible = False
    Me.btnRun.Enabled = True
Else
    Me.lblErrStartChart2.Visible = True
    Me.btnRun.Enabled = False
End If
End Sub
Private Sub txtEndDateChart2_AfterUpdate()
Dim sDate As String
sDate = Trim(Me.txtEndDateChart2.Value)
'Show/hide error label
If VerifyDate(sDate) Then
    Me.lblErrEndChart2.Visible = False
    Me.btnRun.Enabled = True
Else
    Me.lblErrEndChart2.Visible = True
    Me.btnRun.Enabled = False
End If
End Sub


' --  INITIALIZE --
'===================
Private Sub UserForm_Initialize()
Dim lStaIndex As Long
Dim arStation() As String, sStationName As String

'Get default stations items
marStations() = Split(msStations, ",")

'Add stations to Chart1 & Chart2 list items
For lStaIndex = LBound(marStations) To UBound(marStations)
    'Compile station name
    arStation() = Split(marStations(lStaIndex), ":")
    sStationName = arStation(0) & "-" & arStation(1) & " " & Trim(arStation(2))
    'Add station name to list item
    Me.lsbStationChart1.AddItem sStationName    'Chart1
    Me.lsbStationChart2.AddItem sStationName    'Chart2
    'Default list selection
    Me.lsbStationChart1.Selected(0) = True
    Me.lsbStationChart2.Selected(0) = True
Next lStaIndex

'Chart1 start & end date/ times
Me.txtStartDateChart1.Text = Format(Now() - 31, "yyyy/mm/dd")
Me.txtEndDateChart1.Text = Format(Now() - 1, "yyyy/mm/dd")

'Chart2 start & end date/ times
Me.txtStartDateChart2.Text = Format(Now() - 31, "yyyy/mm/dd")
Me.txtEndDateChart2.Text = Format(Now() - 1, "yyyy/mm/dd")

End Sub

' -- FUNCTIONS --
'====================
Private Function VerifyDate(ByVal thisDate As String) As Boolean
Dim bValid As Boolean
bValid = True
'Do nothing if field is blank
If thisDate = "" Then GoTo QuickExit

'Show error label if count not equal 10
If Len(thisDate) = 10 Then
    'Error if "/" not in correct places
    If Not (Mid(thisDate, 5, 1) = "/") Or _
            Not (Mid(thisDate, 5, 1) = "-") Or _
            Not (Mid(thisDate, 8, 1) = "/") Or _
            Not (Mid(thisDate, 8, 1) = "-") Or _
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

