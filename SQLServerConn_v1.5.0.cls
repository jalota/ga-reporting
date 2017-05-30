VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLServerConn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const bDebug As Boolean = True

Private mConn As ADODB.Connection

'-- Class Initialize --
'=======================
Private Sub Class_Initialize()
If Not bDebug Then On Error GoTo ErrHandle
Set mConn = New ADODB.Connection
Exit Sub

ErrHandle:
ErrorHandler lErrNo:=Err.Number, sErrDesc:=Err.Description
Resume Next
End Sub

Sub ConnOpen()
If Not bDebug Then On Error GoTo ErrHandle
mConn.CommandTimeout = 30
mConn.Open ConnString
Exit Sub

ErrHandle:
ErrorHandler lErrNo:=Err.Number, sErrDesc:=Err.Description
Resume Next
End Sub

Sub ConnClose()
On Error Resume Next
If Not mConn Is Nothing Then mConn.Close
Set mConn = Nothing
End Sub

Function Conn() As Connection
Set Conn = mConn
End Function

Function ConnTest() As Boolean
On Error Resume Next
ConnTest = True

mConn.Open ConnString

Clear:
On Error Resume Next
mConn.Close
Set mConn = Nothing
Exit Function

Err:
ConnTest = False
GoTo Clear
End Function

Private Function ConnString() As String
Dim sServer, sPasswd, sUsr, sDb
'SQL Server connection string parameters
If bDebug Then
    sServer = "JL-TDM01"
    sUsr = "sa"
    sPasswd = "Cataly5t"
    sDb = "TDM"
Else
    sServer = "10.121.211.198"
    sUsr = "TDM-Report"
    sPasswd = "Cataly5t"
    sDb = "TDM"
End If
'Compile connection string
ConnString = "Provider=SQLOLEDB" & _
            ";Data Source=" & sServer & _
            ";Persist Security Info=True" & _
            ";User ID=" & sUsr & _
            ";Password=" & sPasswd & _
            ";Initial Catalog=" & sDb
End Function

Private Sub ErrorHandler(lErrNo As Long, sErrDesc As String)
On Error Resume Next
mConn.Close
Set mConn = Nothing
MsgBox "Problem connecting to SQL Server" & vbNewLine & _
        "Contact Catalyst Solution Inc. to resolve the problem", _
        vbCritical + vbOKOnly, _
        "Hyundai Reporting Error"
Resume Next
End Sub

