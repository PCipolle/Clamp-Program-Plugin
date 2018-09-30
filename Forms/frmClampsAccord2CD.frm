VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmClampsAccord2CD 
   Caption         =   "Spacer Selection"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5160
   OleObjectBlob   =   "frmClampsAccord2CD.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmClampsAccord2CD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ImgASpacer_Click()
Dim drw As Drawing
Set drw = App.ActiveDrawing
Dim location As String
Dim frm As Frame
Set frm = App.Frame

frmClampsAccord2CD.Hide

location = frm.PathOfThisAddin

drw.InsertDrawing location + "\layouts\Accord2_CD_A_Spacer.ard", 0, 0, 0
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

Private Sub ImgBSpacer_Click()
Dim drw As Drawing
Set drw = App.ActiveDrawing
Dim location As String
Dim frm As Frame
Set frm = App.Frame

frmClampsAccord2CD.Hide

location = frm.PathOfThisAddin

drw.InsertDrawing location + "\layouts\Accord2_CD_B_Spacer.ard", 0, 0, 0
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

Private Sub ImgCSpacer_Click()
Dim drw As Drawing
Set drw = App.ActiveDrawing
Dim location As String
Dim frm As Frame
Set frm = App.Frame

frmClampsAccord2CD.Hide

location = frm.PathOfThisAddin

drw.InsertDrawing location + "\layouts\Accord2_CD_C_Spacer.ard", 0, 0, 0
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

Private Sub ImgDSpacer_Click()
Dim drw As Drawing
Set drw = App.ActiveDrawing
Dim location As String
Dim frm As Frame
Set frm = App.Frame

frmClampsAccord2CD.Hide

location = frm.PathOfThisAddin

drw.InsertDrawing location + "\layouts\Accord2_CD_D_Spacer.ard", 0, 0, 0
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

Private Sub ImgNoSpacer_Click()
Dim drw As Drawing
Set drw = App.ActiveDrawing
Dim location As String
Dim frm As Frame
Set frm = App.Frame

frmClampsAccord2CD.Hide

location = frm.PathOfThisAddin

drw.InsertDrawing location + "\layouts\Accord2_CD_No_Spacer.ard", 0, 0, 0
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
