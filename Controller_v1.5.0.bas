Attribute VB_Name = "Controller"
Option Explicit

'Globals Vars
'--------------
Private Const msVersion As String = "v1.5.0"
Private Const msReleased As String = "30-May-2017"
Private Const msDeveloper As String = "Jason Lotz <jlotz@catalystsi.com>"

Public Const gsAPP_NAME As String = "Hyundai-Reporting"

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

'-- CONTROLS --
'===============
Sub ShowControlPanel(control As IRibbonControl)
form_ReportControl.Show
End Sub

Sub ShowDashboard(control As IRibbonControl)
Dashboard.Make_Dashboard
End Sub

Sub ShowToolsCharts(control As IRibbonControl)
form_ToolCharts.Show
End Sub

Sub ShowAbout(control As IRibbonControl)
form_About.Show
End Sub


'-- GETTERS --
'==============
Function GetDefStations() As String
GetDefStations = msStations
End Function
Function GetDefLines() As String
GetDefLines = msLines
End Function
Function GetVersion() As String
GetVersion = msVersion
End Function
Function GetDeveloper() As String
GetDeveloper = msDeveloper
End Function
Function GetReleased() As String
GetReleased = msReleased
End Function
Function GetChanges() As String
Dim sChanges As String
sChanges = "## CHANGES LOG" & _
        ", -----------------------" & _
        ",* Hide Control Panel after output report" & _
        ",* Select 'ALL STATIONS' unselects other stations" & _
        ",* Format station report output" & _
        ",* Fix trending period text fields" & _
        ",* Real time reporting"
GetChanges = sChanges
End Function

