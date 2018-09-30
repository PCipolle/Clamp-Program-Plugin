Attribute VB_Name = "Events"
 Function InitAlphacamAddIn(AcamVersion As Long) As Integer
  Dim frm As Frame
  Set frm = App.Frame
  Dim btnId As Long
  Dim location As String
  
  location = frm.PathOfThisAddin
  
  btnId = frm.CreateButtonBar("SCM Clamps")
  frm.AddMenuItem2 "&Insert Clamps", "insertClamps", acamMenuNEW, "SCM Clamps"
      frm.AddButton btnId, location + "\bmp\Insert_Clamps.png", frm.LastMenuCommandID
    
  frm.AddMenuItem2 "&Store Clamps", "storeClamps", acamMenuNEW, "SCM Clamps"
      frm.AddButton btnId, location + "\bmp\Store_Clamps.png", frm.LastMenuCommandID
      
  frm.AddMenuItem2 "&Play/Edit Clamps", "playClamps", acamMenuNEW, "SCM Clamps"
      frm.AddButton btnId, location + "\bmp\Play_Clamps.png", frm.LastMenuCommandID

  InitAlphacamAddIn = 0

 End Function
 Public Sub AfterSaveFile(Filename As String)
 
 App.SelectPost (App.LicomdirPath & "LicomDir\VBMacros\StartUp\XilogPlusPost" & "\XilogPlus.arb")
 
 
 End Sub
  Public Sub AfterCreateNC()
 
 App.SelectPost (App.LicomdirPath & "LicomDir\VBMacros\StartUp\XilogPlusPost" & "\XilogPlus.arb")
 
 
 End Sub
Public Function OnUpdateinsertClamps()
Dim drw As Drawing
Set drw = App.ActiveDrawing

If drw.Clamps.count > 0 Then
    OnUpdateinsertClamps = 0
Else
    OnUpdateinsertClamps = 1
End If

End Function
Public Function OnUpdatestoreClamps()
Dim drw As Drawing
Set drw = App.ActiveDrawing

If drw.Clamps.count = 0 Or drw.Clamps.count = 18 Then
    OnUpdatestoreClamps = 0
Else
    OnUpdatestoreClamps = 1
End If

End Function
Public Function OnUpdateplayClamps()
Dim drw As Drawing
Set drw = App.ActiveDrawing

If drw.Clamps.count = 0 Or drw.Clamps.count = 18 Then
    OnUpdateplayClamps = 0
ElseIf drw.Clamps.Item(1).GetAttributeName(1) <> "BddwUSApcClampsModule" Then
    OnUpdateplayClamps = 0
Else
    OnUpdateplayClamps = 1
End If

End Function

Public Sub storeClamps()

    frmOpNumber.chkMoveClosed.Value = False
    
    frmOpNumber.Show
    
End Sub

Public Sub AfterOpenFile(Filename As String)
Dim drw As Drawing
Set drw = App.ActiveDrawing

If drw.Clamps.count = 0 Then
    Exit Sub
Else
    currentPositions
End If

End Sub
Public Sub GeometriesUpdated()
    If App.ActiveDrawing.Clamps.count = 0 Then
        Exit Sub
    Else
    checkAndMove
    End If
    
End Sub
Public Sub playClamps()
    positionClampsFromData
End Sub
Public Sub insertClamps()
    machineSelection
End Sub


