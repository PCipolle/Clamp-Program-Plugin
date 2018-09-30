VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOpNumber 
   Caption         =   "Get Op Number"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3285
   OleObjectBlob   =   "frmOpNumber.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmOpNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub cmdEnter_Click()
Dim state As Integer
Dim opNum As Integer
opNum = 0

If chkMoveClosed.Value = True Then
    state = 2
Else
    state = 1
End If

If frmOpNumber.txtOpNum = "" Then
    MsgBox "Please enter an operation number"
    Exit Sub

ElseIf frmOpNumber.txtOpNum > 0 Then
    opNum = frmOpNumber.txtOpNum
    writePositionsToAttribute opNum, state
    Exit Sub
ElseIf frmOpNumber.txtOpNum = 0 Then
    MsgBox "Must be greater than zero"
    Exit Sub
ElseIf IsNumeric(frmOpNumber.txtOpNum) = False Then
    MsgBox "Must be a number"
End If


End Sub

Private Sub UserForm_Activate()
frmOpNumber.txtOpNum = ""
End Sub

Private Sub UserForm_Click()

End Sub
