VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_About 
   Caption         =   "About"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6690
   OleObjectBlob   =   "form_About_v1.5.0.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "form_About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private msChanges() As String

'-- Close form --
'=================
Private Sub lblLogo_Click()
Unload Me
End Sub
Private Sub lblReleased_Click()
Unload Me
End Sub
Private Sub txtDeveloper_Click()
Unload Me
End Sub
Private Sub txtReleased_Click()
Unload Me
End Sub
Private Sub lblDeveloper_Click()
Unload Me
End Sub
Private Sub UserForm_Click()
Unload Me
End Sub

'-- Show Changes Log --
'======================
Private Sub lblVersion_Click()
If Not Me.lsbChanges.Visible Then
    Call ShowChangesLog
Else
    Call HideChangesLog
End If
End Sub
Private Sub txtVersion_Click()
If Not Me.lsbChanges.Visible Then
    Call ShowChangesLog
Else
    Call HideChangesLog
End If
End Sub

Private Sub ShowChangesLog()
Dim i As Integer

If Me.lsbChanges.ListCount > 0 Then GoTo Show

Me.lsbChanges.Clear
msChanges() = Split(Controller.GetChanges(), ",")

For i = LBound(msChanges) To UBound(msChanges)
    Me.lsbChanges.AddItem msChanges(i)
Next i

Show:
Me.lsbChanges.Visible = True

End Sub
Private Sub HideChangesLog()
Me.lsbChanges.Visible = False
End Sub


'-- Initialization --
'======================
Private Sub UserForm_Initialize()
With Me
 .txtDeveloper.Caption = Controller.GetDeveloper
 .txtReleased.Caption = Controller.GetReleased
 .txtVersion.Caption = Controller.GetVersion
 
 .lsbChanges.Visible = True
End With
End Sub
