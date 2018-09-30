VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPlayClamps 
   Caption         =   "Play Clamps"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3330
   OleObjectBlob   =   "frmPlayClamps.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPlayClamps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdChangeOp_Click()
Dim temp As String
Dim opNum As Integer
Dim newOpNum As Integer

temp = frmPlayClamps.lblOpNum
If Len(temp) = 5 Then
    opNum = Right(temp, 2)
Else
    opNum = Right(temp, 1)
End If

    changeOperationNumber opNum
    
End Sub

Private Sub cmdPlayBackward_Click()
Dim direction As Integer
    direction = 0
    moveThroughPositions direction
    
End Sub


Private Sub cmdPlayForward_Click()
Dim direction As Integer
    direction = 1
    moveThroughPositions direction

End Sub

Private Sub cmdRemoveOp_Click()
Dim temp As String
Dim opNum As Integer

Dim x As Integer

temp = frmPlayClamps.lblOpNum
If Len(temp) = 5 Then
    opNum = Right(temp, 2)
Else
    opNum = Right(temp, 1)
End If


x = MsgBox("Remove clamp movement at Op " & opNum, vbOKCancel)
If x = 1 Then
    removeClampOperation opNum
Else
    Exit Sub
End If

End Sub

Private Sub UserForm_Terminate()

opsVisible

End Sub
