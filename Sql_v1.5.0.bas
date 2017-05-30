Attribute VB_Name = "Sql"


Function StationReport() As String
Dim sSql As String
Dim sStart As String, sEnd As String

'Define Start & End period markers
sStart = GetStartDate() & " " & GetStartTime()
sEnd = GetEndDate() & " " & GetEndTime()

'Compile sql script
sSql = sSql & "DECLARE @StartPeriod datetime, @EndPeriod datetime " & vbNewLine
sSql = sSql & "SET @StartPeriod = '" & sStart & "' " & vbNewLine
sSql = sSql & "SET @EndPeriod = '" & sEnd & "'; " & vbNewLine
sSql = sSql & "WITH TorqFail(StationObjId,Line,Station,FailCnt) AS ( " & vbNewLine
sSql = sSql & " SELECT st.ObjId AS StationObjId, (LEFT(st.Station,2)) AS Line, st.Station AS StationName, ISNULL(v_TorqFails.FailCnt, 0) AS FailCnt " & vbNewLine
sSql = sSql & " FROM STATION AS st " & vbNewLine
sSql = sSql & " INNER JOIN (  " & vbNewLine
sSql = sSql & "     SELECT ord.StationObjId, COUNT(ord.EventType) AS FailCnt " & vbNewLine
sSql = sSql & "     FROM OrderTraceData AS ord " & vbNewLine
sSql = sSql & "     WHERE ord.EventType = 305 AND ord.RecordTime BETWEEN @StartPeriod AND @EndPeriod " & vbNewLine
sSql = sSql & "     GROUP BY ord.StationObjId, ord.EventType " & vbNewLine
sSql = sSql & " ) AS v_TorqFails  ON st.ObjId = v_TorqFails.StationObjId ) " & vbNewLine
sSql = sSql & ", VehicleClosed(StationObjId,Line,Station,ClosedCnt) AS ( " & vbNewLine
sSql = sSql & " SELECT DISTINCT  st.ObjId AS StationObjId, (LEFT(st.Station,2)) AS Line, st.Station AS StationName,ISNULL(v_StaClosedVeh.StationClosedCnt,0) AS ClosedOrderCnt " & vbNewLine
sSql = sSql & " FROM STATION AS st " & vbNewLine
sSql = sSql & " INNER JOIN ( " & vbNewLine
sSql = sSql & "     SELECT StationObjId, COUNT(Val) AS StationClosedCnt " & vbNewLine
sSql = sSql & "     FROM( " & vbNewLine
sSql = sSql & "         SELECT StationObjId, COUNT(OrderId) AS Val " & vbNewLine
sSql = sSql & "         FROM [OrderTraceData] " & vbNewLine
sSql = sSql & "         WHERE EventType IN (302, 306) " & vbNewLine
sSql = sSql & "         AND RecordTime BETWEEN @StartPeriod AND @EndPeriod " & vbNewLine
sSql = sSql & "         GROUP BY StationObjId, OrderId ) T  " & vbNewLine
sSql = sSql & "     WHERE Val = 1 " & vbNewLine
sSql = sSql & "     GROUP BY StationObjId, Val " & vbNewLine
sSql = sSql & " ) AS v_StaClosedVeh ON st.ObjId = v_StaClosedVeh.StationObjId ) " & vbNewLine
sSql = sSql & ", StationScans(StationObjId,Line,Station,ScanCnt) AS ( " & vbNewLine
sSql = sSql & " SELECT st.ObjId AS StationObjId, (LEFT(st.Station,2)) AS Line, st.Station AS StationName, ISNULL(v_LineProc.VehicleCnt, 0) AS ScanCnt " & vbNewLine
sSql = sSql & " FROM STATION st " & vbNewLine
sSql = sSql & " INNER JOIN ( " & vbNewLine
sSql = sSql & "     SELECT ord.StationObjId, COUNT(ord.EventType) AS VehicleCnt " & vbNewLine
sSql = sSql & "     FROM OrderTraceData AS ord  " & vbNewLine
sSql = sSql & "     WHERE EventType = 303 AND ord.RecordTime BETWEEN @StartPeriod AND @EndPeriod " & vbNewLine
sSql = sSql & "     GROUP BY ord.StationObjId, ord.EventType  " & vbNewLine
sSql = sSql & " ) AS v_LineProc ON st.ObjId = v_LineProc.StationObjId ) " & vbNewLine
sSql = sSql & ", LineScans(StationObjId,Line,Station,ScanCnt) AS ( " & vbNewLine
sSql = sSql & " SELECT st.ObjId AS StationObjId, (LEFT(st.Station,2)) AS Line, st.Station AS StationName, ISNULL(v_LineProc.VehicleCnt, 0) AS ScanCnt " & vbNewLine
sSql = sSql & " FROM STATION st " & vbNewLine
sSql = sSql & " INNER JOIN ( " & vbNewLine
sSql = sSql & "     SELECT ord.StationObjId, COUNT(ord.EventType) AS VehicleCnt " & vbNewLine
sSql = sSql & "     FROM OrderTraceData AS ord  " & vbNewLine
sSql = sSql & "     INNER JOIN STATION AS sta ON ord.StationObjId = sta.ObjId " & vbNewLine
sSql = sSql & "     WHERE EventType = 303 AND sta.StationTypeId = 5 AND ord.RecordTime BETWEEN @StartPeriod AND @EndPeriod " & vbNewLine
sSql = sSql & "     GROUP BY ord.StationObjId, ord.EventType " & vbNewLine
sSql = sSql & " ) AS v_LineProc ON st.ObjId = v_LineProc.StationObjId ) " & vbNewLine
sSql = sSql & ", NGBreakout(StationObjId,NGCnt,TorqHiCnt,TorqLoCnt,AngleHiCnt,AngleLoCnt) AS ( " & vbNewLine
sSql = sSql & " SELECT t.StationObjId, MAX(t.Status) 'Status', MAX(t.TorqHi) 'TorqHigh', MAX(t.TorqLo) 'TorqLow' " & vbNewLine
sSql = sSql & "     , MAX(t.AngleHi) 'AngleHigh', MAX(t.AngleLo) 'AngleLow' " & vbNewLine
sSql = sSql & " FROM ( " & vbNewLine
sSql = sSql & "     SELECT DISTINCT ot.StationObjId " & vbNewLine
sSql = sSql & "         , COUNT(CASE WHEN esp.NAME = 'Status' AND esp.VALUE = 'NOK' THEN esp.VALUE END) 'Status' " & vbNewLine
sSql = sSql & "         , COUNT(CASE WHEN esp.NAME = 'TorqueStatus' AND esp.VALUE = 'High' THEN esp.VALUE END) 'TorqHi' " & vbNewLine
sSql = sSql & "         , COUNT(CASE WHEN esp.NAME = 'TorqueStatus' AND esp.VALUE = 'Low' THEN esp.VALUE END) 'TorqLo' " & vbNewLine
sSql = sSql & "         , COUNT(CASE WHEN esp.NAME = 'AngleStatus' AND esp.VALUE = 'High' THEN esp.VALUE END) 'AngleHi' " & vbNewLine
sSql = sSql & "         , COUNT(CASE WHEN esp.NAME = 'AngleStatus' AND esp.VALUE = 'Low' THEN esp.VALUE END) 'AngleLo' " & vbNewLine
sSql = sSql & "     FROM OrderTraceData AS ot WITH (nolock)  " & vbNewLine
sSql = sSql & "      INNER JOIN OrderTraceDataDetails AS otd ON ot.ID = otd.OrderTraceDataID " & vbNewLine
sSql = sSql & "      INNER JOIN EOR AS e WITH (nolock) ON otd.EORID = e.ID " & vbNewLine
sSql = sSql & "      INNER JOIN EORSpindle AS es ON e.ID = es.EORID " & vbNewLine
sSql = sSql & "      INNER JOIN EORSpindleParam AS esp ON es.ID = esp.SPINDLEID " & vbNewLine
sSql = sSql & "      INNER JOIN STATION AS st ON ot.StationObjId = st.ObjId " & vbNewLine
sSql = sSql & "     WHERE e.RUNDOWNTIME BETWEEN @StartPeriod AND @EndPeriod " & vbNewLine
sSql = sSql & "     GROUP BY ot.StationObjId, esp.NAME ) t " & vbNewLine
sSql = sSql & " GROUP BY t.StationObjId ) " & vbNewLine
sSql = sSql & "SELECT Station, PeriodStart, PeriodEnd, TotalVehicle, TotalPass, TotalScans " & vbNewLine
sSql = sSql & " , CASE WHEN TotalPass > 0 THEN CAST(CAST(((TotalPass / CAST(TotalVehicle AS DECIMAL)) * 100) AS INT) AS VARCHAR) ELSE '0' END AS PassRatio " & vbNewLine
sSql = sSql & " , CASE WHEN TotalScans > 0 THEN CAST(CAST(((TotalScans / CAST(TotalVehicle AS DECIMAL)) * 100) AS INT) AS VARCHAR) ELSE '0' END AS ScanRatio " & vbNewLine
sSql = sSql & " , TotalNG, TorqueHi, TorqueLo, AngleHi, AngleLo " & vbNewLine
sSql = sSql & "FROM ( " & vbNewLine
sSql = sSql & " SELECT st.ObjId AS StationObjId, (LEFT(st.Station,2)) AS Line, @StartPeriod AS PeriodStart, st.Station, @EndPeriod AS PeriodEnd, ISNULL(LineScans.ScanCnt,0) AS TotalVehicle " & vbNewLine
sSql = sSql & "     , ISNULL(StationScans.ScanCnt,0) AS TotalScans, ISNULL(VehicleClosed.ClosedCnt,0) AS TotalPass " & vbNewLine
sSql = sSql & "     , ISNULL(ng.NGCnt,0) 'TotalNG', ISNULL(ng.TorqHiCnt,0) 'TorqueHi', ISNULL(ng.TorqLoCnt,0) 'TorqueLo', ISNULL(ng.AngleHiCnt,0) 'AngleHi', ISNULL(ng.AngleLoCnt,0) 'AngleLo', st.WorkCenterObjId " & vbNewLine
sSql = sSql & " FROM STATION AS st " & vbNewLine
sSql = sSql & " INNER JOIN WORKCENTER AS wc ON st.WorkCenterObjId = wc.ObjId"
sSql = sSql & " LEFT OUTER JOIN LineScans ON LEFT(st.Station,2) = LineScans.Line " & vbNewLine
sSql = sSql & " LEFT OUTER JOIN StationScans ON st.ObjId = StationScans.StationObjId " & vbNewLine
sSql = sSql & " LEFT OUTER JOIN VehicleClosed ON st.ObjId = VehicleClosed.StationObjId " & vbNewLine
sSql = sSql & " LEFT OUTER JOIN TorqFail ON st.ObjId = TorqFail.StationObjId " & vbNewLine
sSql = sSql & " LEFT OUTER JOIN NGBreakout AS ng ON st.ObjId = ng.StationObjId ) AS rpt " & vbNewLine

StationReport = sSql
End Function

Function LineReport() As String
Dim sSql As String
Dim sStart As String, sEnd As String

'Define Start & End period markers
sStart = GetStartDate() & " " & GetStartTime()
sEnd = GetEndDate() & " " & GetEndTime()

'Compile sql script
sSql = "DECLARE @StartPeriod datetime, @EndPeriod datetime " & vbNewLine
sSql = sSql & "SET @StartPeriod = '" & sStart & "' " & vbNewLine
sSql = sSql & "SET @EndPeriod = '" & sEnd & "'; " & vbNewLine
sSql = sSql & "WITH LinePass(WorkCenterObjId,PassCnt) AS ( " & vbNewLine
sSql = sSql & " SELECT st.WorkCenterObjId,ISNULL(v_ClosedVehicle.StationClosedCnt,0) AS PassCnt " & vbNewLine
sSql = sSql & " FROM STATION AS st " & vbNewLine
sSql = sSql & " INNER JOIN ( " & vbNewLine
sSql = sSql & "     SELECT DISTINCT StationObjId, COUNT(Val) AS StationClosedCnt " & vbNewLine
sSql = sSql & "     FROM( SELECT StationObjId, COUNT(OrderId) AS Val " & vbNewLine
sSql = sSql & "         FROM [OrderTraceData] " & vbNewLine
sSql = sSql & "         WHERE EventType IN (302, 306) AND RecordTime BETWEEN @StartPeriod AND @EndPeriod " & vbNewLine
sSql = sSql & "         GROUP BY StationObjId, OrderId ) T  " & vbNewLine
sSql = sSql & "     WHERE Val = 1 " & vbNewLine
sSql = sSql & "     GROUP BY StationObjId, Val " & vbNewLine
sSql = sSql & " ) AS v_ClosedVehicle ON st.ObjId = v_ClosedVehicle.StationObjId ) " & vbNewLine
sSql = sSql & ", LineScan(WorkCenterObjId,ScanCnt) AS ( " & vbNewLine
sSql = sSql & " SELECT DISTINCT st.WorkCenterObjId, COUNT(ord.EventType) AS ScanCnt " & vbNewLine
sSql = sSql & " FROM OrderTraceData AS ord  " & vbNewLine
sSql = sSql & "  INNER JOIN STATION AS st ON ord.StationObjId = st.ObjId " & vbNewLine
sSql = sSql & " WHERE ord.EventType = 303 " & vbNewLine
sSql = sSql & " AND ord.RecordTime BETWEEN @StartPeriod AND @EndPeriod " & vbNewLine
sSql = sSql & " GROUP BY st.WorkCenterObjId, st.ObjId, ord.EventType ) " & vbNewLine
sSql = sSql & ", TotalScans(WorkCenterObjId,ScanCnt) AS ( " & vbNewLine
sSql = sSql & " SELECT st.WorkCenterObjId, ISNULL(v_LineProc.VehicleCnt, 0) AS ScanCnt " & vbNewLine
sSql = sSql & " FROM STATION st " & vbNewLine
sSql = sSql & " INNER JOIN (  " & vbNewLine
sSql = sSql & "     SELECT ord.StationObjId, COUNT(ord.EventType) AS VehicleCnt " & vbNewLine
sSql = sSql & "     FROM OrderTraceData AS ord  " & vbNewLine
sSql = sSql & "      INNER JOIN STATION AS sta ON ord.StationObjId = sta.ObjId " & vbNewLine
sSql = sSql & "     WHERE EventType = 303  " & vbNewLine
sSql = sSql & "     AND sta.StationTypeId = 5 " & vbNewLine
sSql = sSql & "     AND ord.RecordTime BETWEEN @StartPeriod AND @EndPeriod " & vbNewLine
sSql = sSql & "     GROUP BY ord.StationObjId, ord.EventType " & vbNewLine
sSql = sSql & " ) AS v_LineProc ON st.ObjId = v_LineProc.StationObjId ) " & vbNewLine
sSql = sSql & ", NGBreakout(WorkCenterObjId,NGCnt,TorqHiCnt,TorqLoCnt,AngleHiCnt,AngleLoCnt) AS ( " & vbNewLine
sSql = sSql & " SELECT t.WorkCenterObjId, MAX(t.Status) 'Status', MAX(t.TorqHi) 'TorqHigh', MAX(t.TorqLo) 'TorqLow', MAX(t.AngleHi) 'AngleHigh', MAX(t.AngleLo) 'AngleLow' " & vbNewLine
sSql = sSql & " FROM (  " & vbNewLine
sSql = sSql & "     SELECT st.WorkCenterObjId, COUNT(CASE WHEN esp.NAME = 'Status' AND esp.VALUE = 'NOK' THEN esp.VALUE END) 'Status' " & vbNewLine
sSql = sSql & "         , COUNT(CASE WHEN esp.NAME = 'TorqueStatus' AND esp.VALUE = 'High' THEN esp.VALUE END) 'TorqHi', COUNT(CASE WHEN esp.NAME = 'TorqueStatus' AND esp.VALUE = 'Low' THEN esp.VALUE END) 'TorqLo' " & vbNewLine
sSql = sSql & "         , COUNT(CASE WHEN esp.NAME = 'AngleStatus' AND esp.VALUE = 'High' THEN esp.VALUE END) 'AngleHi', COUNT(CASE WHEN esp.NAME = 'AngleStatus' AND esp.VALUE = 'Low' THEN esp.VALUE END) 'AngleLo' " & vbNewLine
sSql = sSql & "     FROM OrderTraceData AS ot WITH (nolock)  " & vbNewLine
sSql = sSql & "       INNER JOIN OrderTraceDataDetails AS otd ON ot.ID = otd.OrderTraceDataID " & vbNewLine
sSql = sSql & "       INNER JOIN EOR AS e WITH (nolock) ON otd.EORID = e.ID " & vbNewLine
sSql = sSql & "       INNER JOIN EORSpindle AS es ON e.ID = es.EORID " & vbNewLine
sSql = sSql & "       INNER JOIN EORSpindleParam AS esp ON es.ID = esp.SPINDLEID " & vbNewLine
sSql = sSql & "       INNER JOIN STATION AS st ON ot.StationObjId = st.ObjId " & vbNewLine
sSql = sSql & "     WHERE e.RUNDOWNTIME BETWEEN @StartPeriod AND @EndPeriod " & vbNewLine
sSql = sSql & "     GROUP BY st.WorkCenterObjId, esp.NAME ) t " & vbNewLine
sSql = sSql & " GROUP BY t.WorkCenterObjId ) " & vbNewLine
sSql = sSql & "SELECT WorkCenter, @StartPeriod 'PeriodStart', @EndPeriod 'PeriodEnd' " & vbNewLine
sSql = sSql & " , CASE WHEN TotalScans > 0 THEN CAST(CAST(((HighPass / CAST(TotalScans AS DECIMAL)) * 100) AS INT) AS VARCHAR) ELSE '0' END AS 'HighPassRatio' " & vbNewLine
sSql = sSql & " , CASE WHEN TotalScans > 0 THEN CAST(CAST(((LowPass / CAST(TotalScans AS DECIMAL)) * 100) AS INT) AS VARCHAR) ELSE '0' END AS 'LowPassRatio' " & vbNewLine
sSql = sSql & " , CASE WHEN TotalScans > 0 THEN CAST(CAST(((AvgPass / CAST(TotalScans AS DECIMAL)) * 100) AS INT) AS VARCHAR) ELSE '0' END AS 'AvgPassRatio' " & vbNewLine
sSql = sSql & " , CASE WHEN TotalScans > 0 THEN CAST(CAST(((HighScan / CAST(TotalScans AS DECIMAL)) * 100) AS INT) AS VARCHAR) ELSE '0' END AS 'HighScanRatio' " & vbNewLine
sSql = sSql & " , CASE WHEN TotalScans > 0 THEN CAST(CAST(((LowScan / CAST(TotalScans AS DECIMAL)) * 100) AS INT) AS VARCHAR) ELSE '0' END AS 'LowScanRatio' " & vbNewLine
sSql = sSql & " , CASE WHEN TotalScans > 0 THEN CAST(CAST(((AvgScan / CAST(TotalScans AS DECIMAL)) * 100) AS INT) AS VARCHAR) ELSE '0' END AS 'AvgScanRatio' " & vbNewLine
sSql = sSql & " , NG_Cnt, TorqHi_Cnt, TorqLo_Cnt, AngleHi_Cnt, AngleLo_Cnt " & vbNewLine
sSql = sSql & "FROM ( " & vbNewLine
sSql = sSql & " SELECT DISTINCT wc.ObjId AS WorkCenterObjId, wc.WorkCenter, MAX(lc.ScanCnt) OVER (PARTITION BY wc.WorkCenter) AS HighScan, MIN(lc.ScanCnt) OVER (PARTITION BY wc.WorkCenter) AS LowScan " & vbNewLine
sSql = sSql & "     , AVG(lc.ScanCnt) OVER (PARTITION BY wc.WorkCenter) AS AvgScan, MAX(lp.PassCnt) OVER (PARTITION BY wc.WorkCenter) AS HighPass, MIN(lp.PassCnt) OVER (PARTITION BY wc.WorkCenter) AS LowPass " & vbNewLine
sSql = sSql & "     , AVG(lp.PassCnt) OVER (PARTITION BY wc.WorkCenter) AS AvgPass, ISNULL(ts.ScanCnt,0) AS TotalScans, ISNULL(ng.NGCnt,0) AS NG_Cnt, ISNULL(ng.TorqHiCnt,0) AS TorqHi_Cnt " & vbNewLine
sSql = sSql & "     , ISNULL(ng.TorqLoCnt,0) AS TorqLo_Cnt, ISNULL(ng.AngleHiCnt,0) AS AngleHi_Cnt, ISNULL(ng.AngleLoCnt,0) AS AngleLo_Cnt " & vbNewLine
sSql = sSql & " FROM WORKCENTER AS wc " & vbNewLine
sSql = sSql & "  LEFT OUTER JOIN LinePass AS lp ON wc.ObjId = lp.WorkCenterObjId " & vbNewLine
sSql = sSql & "  LEFT OUTER JOIN LineScan AS lc ON wc.ObjId = lc.WorkCenterObjId " & vbNewLine
sSql = sSql & "  LEFT OUTER JOIN TotalScans AS ts ON wc.ObjId = ts.WorkCenterObjId " & vbNewLine
sSql = sSql & "  LEFT OUTER JOIN NGBreakout AS ng ON wc.ObjId = ng.WorkCenterObjId ) t  " & vbNewLine

LineReport = sSql
End Function

Function FacilityMonthly() As String
Dim s As String
s = "SELECT [MonthName] ' ', [Pass], [Scan] " & vbNewLine
s = s & "FROM [Report_FacilityMonthly] " & vbNewLine
s = s & "WHERE SUBSTRING([MonthDate],1,4) = '" & Format(Now(), "yyyy") & "' " & vbNewLine
s = s & "ORDER BY [MonthDate] "
FacilityMonthly = s
End Function

Function LineStatus() As String
Dim sDate As String
'Set date for last completed shift
If Format(Now(), "ddd") = "Mon" Then
    sDate = Format(Now() - 3, "yyyy-mm-dd")
ElseIf Format(Now(), "ddd") = "Sun" Then
    sDate = Format(Now() - 2, "yyyy-mm-dd")
Else
    sDate = Format(Now() - 1, "yyyy-mm-dd")
End If

Dim s As String
s = "SELECT LTRIM(RTRIM(SUBSTRING(Area,3,LEN(Area)))) 'Line', AVG(PassRatio) 'PassRatio', AVG(ScanRatio) 'ScanRatio' " & vbNewLine
s = s & "FROM ( " & vbNewLine
s = s & "   SELECT AREA.Area, Report_StationDaily.Station, PassRatio, ScanRatio " & vbNewLine
s = s & "   FROM Report_StationDaily " & vbNewLine
s = s & "   INNER JOIN STATION ON Report_StationDaily.Station = STATION.Station " & vbNewLine
s = s & "   INNER JOIN WORKCENTER ON STATION.WorkCenterObjId = WORKCENTER.ObjId " & vbNewLine
s = s & "   INNER JOIN AREA ON WORKCENTER.AreaObjId = AREA.ObjId " & vbNewLine
s = s & "   WHERE ProdDate = '" & sDate & "' ) t " & vbNewLine
s = s & "WHERE PassRatio <> 0 GROUP BY Area ORDER BY Area "
LineStatus = s
End Function

Function DailyStatus(sStart As String, sEnd As String, sShift As String) As String
Dim s As String
'Compile sql string
s = "SELECT LTRIM(RTRIM(SUBSTRING(Facility,3,LEN(Facility)))) 'Line', AVG(PassRatio) 'PassRatio', AVG(ScanRatio) 'ScanRatio' " & vbNewLine
s = s & "FROM ( " & vbNewLine
s = s & "   SELECT FACILITY.Facility, Report_StationDaily.Station, PassRatio, ScanRatio " & vbNewLine
s = s & "   FROM Report_StationDaily " & vbNewLine
s = s & "   INNER JOIN STATION ON Report_StationDaily.Station = STATION.Station " & vbNewLine
s = s & "   INNER JOIN WORKCENTER ON STATION.WorkCenterObjId = WORKCENTER.ObjId " & vbNewLine
s = s & "   INNER JOIN AREA ON WORKCENTER.AreaObjId = AREA.ObjId " & vbNewLine
s = s & "   INNER JOIN FACILITY ON AREA.FacilityObjId = FACILITY.ObjId " & vbNewLine
s = s & "   WHERE ProdDate BETWEEN '" & sStart & "' AND '" & sEnd & "' AND ProdShift LIKE '" & sShift & "' ) t " & vbNewLine
s = s & "WHERE PassRatio <> 0 GROUP BY Facility "
DailyStatus = s
End Function

Function StationFiveShift(sStart As String, sEnd As String, sShift As String) As String
Dim s As String
'Compile sql string
s = "SELECT t.Station, t.ProdDate, ' ', t.TotalVehicle, t.TotalPass, t.TotalScan " & vbNewLine
s = s & "   , CASE WHEN t.TotalVehicle > 0 THEN CAST(CAST(((t.TotalPass / CAST(t.TotalVehicle AS DECIMAL)) * 100) AS INT) AS VARCHAR) ELSE '0' END AS 'PassRatio' " & vbNewLine
s = s & "   , CASE WHEN t.TotalVehicle > 0 THEN CAST(CAST(((t.TotalScan / CAST(t.TotalVehicle AS DECIMAL)) * 100) AS INT) AS VARCHAR) ELSE '0' END AS 'ScanRatio' " & vbNewLine
s = s & "   , t.TotalNG, t.TorqueHigh, t.TorqueLow, t.AngleHigh, t.AngleLow  " & vbNewLine
s = s & "FROM STATION st  " & vbNewLine
s = s & "INNER JOIN (   SELECT DISTINCT rs.ProdDate, rs.Station, SUM(rs.TotalVehicle) OVER (PARTITION BY rs.Station, rs.ProdDate) 'TotalVehicle'  " & vbNewLine
s = s & "       , SUM(rs.TotalPass) OVER (PARTITION BY rs.Station, rs.ProdDate) 'TotalPass', SUM(rs.TotalScans) OVER (PARTITION BY rs.Station, rs.ProdDate) 'TotalScan'  " & vbNewLine
s = s & "       , SUM(rs.TotalNG) OVER (PARTITION BY rs.Station, rs.ProdDate) 'TotalNG', SUM(rs.TorqueHigh) OVER (PARTITION BY rs.Station, rs.ProdDate) 'TorqueHigh'  " & vbNewLine
s = s & "       , SUM(rs.TorqueLow) OVER (PARTITION BY rs.Station, rs.ProdDate) 'TorqueLow', SUM(rs.AngleHigh) OVER (PARTITION BY rs.Station, rs.ProdDate) 'AngleHigh'  " & vbNewLine
s = s & "       , SUM(rs.AngleLow) OVER (PARTITION BY rs.Station, rs.ProdDate) 'AngleLow'  " & vbNewLine
s = s & "   FROM Report_StationDaily rs  " & vbNewLine
s = s & "    INNER JOIN STATION st ON rs.Station = st.Station " & vbNewLine
s = s & "   WHERE rs.ProdDate BETWEEN '" & sStart & "' AND '" & sEnd & "' AND rs.ProdShift LIKE '" & sShift & "'  " & vbNewLine
s = s & ") t ON st.Station = t.Station   " & vbNewLine
StationFiveShift = s
End Function

Function LineFiveShift(sStart As String, sEnd As String, sShift As String) As String
Dim s As String
'Compile sql string
s = " SELECT WorkCenter, PeriodStart, PeriodEnd, HighPassRatio, LowPassRatio, AvgPassRatio, HighScanRatio, LowScanRatio, AvgScanRatio, TotalNG, TorqueHigh, TorqueLow, AngleHigh, AngleLow " & vbNewLine
s = s & " FROM ( " & vbNewLine
s = s & "    SELECT DISTINCT WorkCenterObjId, WorkCenter, ProdDate AS 'PeriodStart', ' ' AS 'PeriodEnd', CASE WHEN TotalVehicle > 0 THEN CAST(CAST(((HighPass / CAST(TotalVehicle AS DECIMAL)) * 100) AS INT) AS VARCHAR) ELSE '0' END 'HighPassRatio'  " & vbNewLine
s = s & "      , CASE WHEN TotalVehicle > 0 THEN CAST(CAST(((LowPass / CAST(TotalVehicle AS DECIMAL)) * 100) AS INT) AS VARCHAR) ELSE '0' END 'LowPassRatio', CASE WHEN TotalVehicle > 0 THEN CAST(CAST(((AvgPass / CAST(TotalVehicle AS DECIMAL)) * 100) AS INT) AS VARCHAR) ELSE '0' END 'AvgPassRatio'  " & vbNewLine
s = s & "      , CASE WHEN TotalVehicle > 0 THEN CAST(CAST(((HighScan / CAST(TotalVehicle AS DECIMAL)) * 100) AS INT) AS VARCHAR) ELSE '0' END 'HighScanRatio', CASE WHEN TotalVehicle > 0 THEN CAST(CAST(((LowScan / CAST(TotalVehicle AS DECIMAL)) * 100) AS INT) AS VARCHAR) ELSE '0' END 'LowScanRatio'  " & vbNewLine
s = s & "      , CASE WHEN TotalVehicle > 0 THEN CAST(CAST(((AvgScan / CAST(TotalVehicle AS DECIMAL)) * 100) AS INT) AS VARCHAR) ELSE '0' END 'AvgScanRatio', TotalNG, TorqueHigh, TorqueLow, AngleHigh, AngleLow  " & vbNewLine
s = s & "   FROM (  " & vbNewLine
s = s & "      SELECT DISTINCT t.ProdDate, wc.WorkCenter, t.WorkCenterObjId, t.TotalVehicle, t.TotalScans, MAX(t.TotalPass) OVER (PARTITION BY t.ProdDate, t.WorkCenterObjId) 'HighPass'  " & vbNewLine
s = s & "          , MIN(t.TotalPass) OVER (PARTITION BY t.ProdDate, t.WorkCenterObjId) 'LowPass', AVG(t.TotalPass) OVER (PARTITION BY t.ProdDate, t.WorkCenterObjId) 'AvgPass', MAX(t.TotalScans) OVER (PARTITION BY t.ProdDate, t.WorkCenterObjId) 'HighScan'  " & vbNewLine
s = s & "          , MIN(t.TotalScans) OVER (PARTITION BY t.ProdDate, t.WorkCenterObjId) 'LowScan', AVG(t.TotalScans) OVER (PARTITION BY t.ProdDate, t.WorkCenterObjId) 'AvgScan', SUM(t.TotalNG) OVER (PARTITION BY t.ProdDate, t.WorkCenterObjId) 'TotalNG'  " & vbNewLine
s = s & "          , SUM(t.TorqueHigh) OVER (PARTITION BY t.ProdDate, t.WorkCenterObjId) 'TorqueHigh', SUM(t.TorqueLow) OVER (PARTITION BY t.ProdDate, t.WorkCenterObjId) 'TorqueLow', SUM(t.AngleHigh) OVER (PARTITION BY t.ProdDate, t.WorkCenterObjId) 'AngleHigh'  " & vbNewLine
s = s & "          , SUM(t.AngleLow) OVER (PARTITION BY t.ProdDate, t.WorkCenterObjId) 'AngleLow'  " & vbNewLine
s = s & "      FROM WORKCENTER wc  " & vbNewLine
s = s & "      INNER JOIN ( SELECT DISTINCT ProdDate, st.WorkCenterObjId, SUM(rs.TotalVehicle) OVER (PARTITION BY rs.ProdDate, rs.Station) 'TotalVehicle', SUM(rs.TotalPass) OVER (PARTITION BY rs.ProdDAte, rs.Station) 'TotalPass'  " & vbNewLine
s = s & "              , SUM(rs.TotalScans) OVER (PARTITION BY rs.ProdDate, rs.Station) 'TotalScans', SUM(rs.TotalNG) OVER (PARTITION BY rs.ProdDate, rs.Station) 'TotalNG', SUM(rs.TorqueHigh) OVER (PARTITION BY rs.ProdDate, rs.Station) 'TorqueHigh'  " & vbNewLine
s = s & "              , SUM(rs.TorqueLow) OVER (PARTITION BY rs.ProdDate, rs.Station) 'TorqueLow', SUM(rs.AngleHigh) OVER (PARTITION BY rs.ProdDate, rs.Station) 'AngleHigh', SUM(rs.AngleLow) OVER (PARTITION BY rs.ProdDate, rs.Station) 'AngleLow'  " & vbNewLine
s = s & "          FROM STATION st  " & vbNewLine
s = s & "          INNER JOIN Report_StationDaily rs ON st.Station = rs.Station  " & vbNewLine
s = s & "          WHERE ProdDate BETWEEN '" & sStart & "' AND '" & sEnd & "' AND ProdShift LIKE '" & sShift & "'  " & vbNewLine
s = s & "      ) t ON wc.ObjId = t.WorkCenterObjId  " & vbNewLine
s = s & "       WHERE TotalPass <> 0 OR TotalScans <> 0  " & vbNewLine
s = s & "   ) x  " & vbNewLine
s = s & " ) y " & vbNewLine
LineFiveShift = s
End Function

Function MultiToolChart(sStart As String, sEnd As String, sStation As String) As String
Dim s As String
s = " SELECT DISTINCT Station " & vbNewLine
s = s & "   , CASE WHEN TotalVehicle > 0 THEN CAST(CAST(((TotalPass / CAST(TotalVehicle AS DECIMAL)) * 100) AS INT) AS VARCHAR) ELSE '0' END AS 'PassRatio' " & vbNewLine
s = s & "   , CASE WHEN TotalVehicle > 0 THEN CAST(CAST(((TotalScans / CAST(TotalVehicle AS DECIMAL)) * 100) AS INT) AS VARCHAR) ELSE '0' END AS 'ScanRatio' " & vbNewLine
s = s & " FROM ( " & vbNewLine
s = s & "   SELECT DISTINCT ProdDate, Station " & vbNewLine
s = s & "       , SUM(TotalVehicle) OVER (PARTITION BY Station) 'TotalVehicle' " & vbNewLine
s = s & "       , SUM(TotalPass) OVER (PARTITION BY Station) 'TotalPass' " & vbNewLine
s = s & "       , SUM(TotalScans) OVER (PARTITION BY Station) 'TotalScans' " & vbNewLine
s = s & "   FROM Report_StationDaily " & vbNewLine
s = s & "   WHERE ProdDate BETWEEN '" & sStart & "' AND '" & sEnd & "' " & vbNewLine
s = s & "   GROUP BY Station, ProdDate, TotalVehicle, TotalPass, TotalScans " & vbNewLine
s = s & "   ) t " & vbNewLine
s = s & "WHERE Station IN (" & sStation & ") " & vbNewLine
MultiToolChart = s
End Function

Function SingleToolChart(sStart As String, sEnd As String, sStation As String) As String
Dim s As String
s = " SELECT DISTINCT ProdDate " & vbNewLine
s = s & "   , CASE WHEN TotalVehicle > 0 THEN CAST(CAST(((TotalPass / CAST(TotalVehicle AS DECIMAL)) * 100) AS INT) AS VARCHAR) ELSE '0' END AS 'PassRatio'"
s = s & "   , CASE WHEN TotalVehicle > 0 THEN CAST(CAST(((TotalScans / CAST(TotalVehicle AS DECIMAL)) * 100) AS INT) AS VARCHAR) ELSE '0' END AS 'ScanRatio' " & vbNewLine
s = s & "FROM ( " & vbNewLine
s = s & "   SELECT DISTINCT ProdDate, Station " & vbNewLine
s = s & "       , SUM(TotalVehicle) OVER (PARTITION BY Station, ProdDate) 'TotalVehicle' " & vbNewLine
s = s & "       , SUM(TotalPass) OVER (PARTITION BY Station, ProdDate) 'TotalPass' " & vbNewLine
s = s & "       , SUM(TotalScans) OVER (PARTITION BY Station, ProdDate) 'TotalScans' " & vbNewLine
s = s & "   FROM Report_StationDaily " & vbNewLine
s = s & "   WHERE ProdDate BETWEEN '" & sStart & "' AND '" & sEnd & "' " & vbNewLine
s = s & "   GROUP BY Station, ProdDate, TotalVehicle, TotalPass, TotalScans " & vbNewLine
s = s & "   ) t " & vbNewLine
s = s & " WHERE Station = '" & sStation & "' "
SingleToolChart = s
End Function
