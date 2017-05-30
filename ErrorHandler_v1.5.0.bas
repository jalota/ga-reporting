Attribute VB_Name = "ErrorHandler"
Option Explicit
Option Private Module

Public Const gbDEBUG_MODE As Boolean = True     ' True enables debug mode, False disables it.
Public Const glHANDLED_ERROR As Long = 9999     ' Run-time error number for our custom errors.
Public Const glCUSTOM_ERROR As Long = 9998
Public Const glUSER_CANCEL As Long = 18         ' The error number generated when the user cancels program execution.

Private Const msSILENT_ERROR As String = "UserCancel"   ' Used by the central error handler to bail out silently on user cancel.
Private Const msFILE_ERROR_LOG As String = "error.log"  ' The name of the file where error messages will be logged to.


Public Function bCentralErrorHandler( _
            ByVal sMODULE As String, _
            ByVal sProc As String, _
            Optional ByVal sFile As String, _
            Optional ByVal bEntryPoint As Boolean) As Boolean

    Static sErrMsg As String
    
    Dim iFile As Integer
    Dim lErrNum As Long
    Dim sFullSource As String
    Dim sPath As String
    Dim sLogText As String
    
    ' Grab the error info before it's cleared by
    ' On Error Resume Next below.
    lErrNum = Err.Number
    ' If this is a user cancel, set the silent error flag
    ' message. This will cause the error to be ignored.
    If lErrNum = glUSER_CANCEL Then sErrMsg = msSILENT_ERROR
    ' If this is the originating error, the static error
    ' message variable will be empty. In that case, store
    ' the originating error message in the static variable.
    If Len(sErrMsg) = 0 Then sErrMsg = Err.Description

    ' We cannot allow errors in the central error handler.
    On Error Resume Next
    
    ' Load the default filename if required.
    If Len(sFile) = 0 Then sFile = ThisWorkbook.Name
    
    ' Get the application directory.
    sPath = ThisWorkbook.Path
    If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"
    
    ' Construct the fully-qualified error source name.
    sFullSource = "[" & sFile & "]" & sMODULE & "." & sProc

    ' Create the error text to be logged.
    sLogText = "  " & sFullSource & ", Error " & _
                        CStr(lErrNum) & ": " & sErrMsg
    
    ' Open the log file, write out the error information and
    ' close the log file.
    iFile = FreeFile()
    Open sPath & msFILE_ERROR_LOG For Append As #iFile
    Print #iFile, Format$(Now(), "mm/dd/yy hh:mm:ss"); sLogText
    If bEntryPoint Then Print #iFile,
    Close #iFile
        
    ' Do not display or debug silent errors.
    If sErrMsg <> msSILENT_ERROR Then
    
        ' Show the error message when we reach the entry point
        ' procedure or immediately if we are in debug mode.
        If bEntryPoint Or gbDEBUG_MODE Then
            Application.ScreenUpdating = True
            MsgBox sErrMsg, vbCritical, gsAPP_NAME
            ' Clear the static error message variable once
            ' we've reached the entry point so that we're ready
            ' to handle the next error.
            sErrMsg = vbNullString
        End If
        
        ' The return vale is the debug mode status.
        bCentralErrorHandler = gbDEBUG_MODE
        
    Else
        ' If this is a silent error, clear the static error
        ' message variable when we reach the entry point.
        If bEntryPoint Then sErrMsg = vbNullString
        bCentralErrorHandler = False
    End If
    
End Function
