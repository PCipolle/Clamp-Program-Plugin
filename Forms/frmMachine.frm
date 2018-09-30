VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMachine 
   Caption         =   "Machine and Field"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5085
   OleObjectBlob   =   "frmMachine.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMachine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub imgAccord1AB_Click()

frmMachine.Hide

frmClampsAccord1AB.Show

End Sub

Private Sub imgAccord1CD_Click()

frmMachine.Hide

frmClampsAccord1CD.Show

End Sub

Private Sub imgAccord25AB_Click()

frmMachine.Hide

Dim drw As Drawing
Set drw = App.ActiveDrawing
Dim location As String
Dim frm As Frame
Set frm = App.Frame

location = frm.PathOfThisAddin

drw.InsertDrawing location + "\layouts\Accord25_Clamp_AB.ard", 0, 0, 0
For i = 1 To drw.Layers.count
    If drw.Layers.Item(i).Attribute("LicomUKSAJFixtureLayer") = 1 Then
        layNum = i
        Exit For
        
    End If
Next i

drw.Layers.Item(layNum).ColorRGB = RGB(0, 113, 225)

drw.Redraw
drw.RedrawShadedViews
drw.Refresh

End Sub

Private Sub imgAccord25CD_Click()

frmMachine.Hide

Dim drw As Drawing
Set drw = App.ActiveDrawing
Dim location As String
Dim frm As Frame
Set frm = App.Frame

location = frm.PathOfThisAddin

drw.InsertDrawing location + "\layouts\Accord25_Clamp_CD.ard", 0, 0, 0
For i = 1 To drw.Layers.count
    If drw.Layers.Item(i).Attribute("LicomUKSAJFixtureLayer") = 1 Then
        layNum = i
        Exit For
        
    End If
Next i

drw.Layers.Item(layNum).ColorRGB = RGB(0, 113, 225)

drw.Redraw
drw.RedrawShadedViews
drw.Refresh

End Sub

Private Sub imgAccord2AB_Click()

frmMachine.Hide

frmClampsAccord2AB.Show

End Sub

Private Sub imgAccord2CD_Click()

frmMachine.Hide

frmClampsAccord2CD.Show


End Sub

Private Sub Label1_Click()

End Sub
