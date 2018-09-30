Attribute VB_Name = "Main"
'Created by: Parker Cipolle
'Version:    3.0
'Date:       December 30, 2015
'This program is currently in testing phase
'Please do not use this program without permission

Const data As String = "BddwUSApcClampsModule"

Dim pos(32) As Double
Dim posFromData() As Double
Dim posFlag As Integer

Public Function RGB(Red As Integer, Green As Integer, Blue As Integer) As OLE_COLOR

If Red < 0 Or Red > 255 Or Green < 0 Or Green > 255 Or Blue < 0 Or Blue > 255 Then

RGB = 0

Exit Function

End If

RGB = Red + Green * &H100& + Blue * &H10000

End Function

Public Sub writePositionsToAttribute(opNum As Integer, state As Integer)
frmOpNumber.Hide
Dim drw As Drawing
Set drw = App.ActiveDrawing
Dim iterate As Integer

If drw.Clamps.count = 30 Or drw.Clamps.count = 45 Then
    iterate = 15
ElseIf drw.Clamps.count = 32 Or drw.Clamps.count = 48 Then
    iterate = 16
Else
    MsgBox "Store positions failed! There must be 15 or 16 clamps in the drawing"
    Exit Sub
End If

For i = 1 To iterate
    
    If drw.Clamps.Item(i).Attribute(data) = "" Then
        drw.Clamps.Item(i).Attribute(data) = "OP" & opNum & "," & Math.Round(drw.Clamps.Item(i).BasePointX, 3) & "," & Math.Round(drw.Clamps.Item(i).BasePointY, 3) & "," & state
    Else
        drw.Clamps.Item(i).Attribute(data) = drw.Clamps.Item(i).Attribute(data) & "," & "OP" & opNum & "," & Math.Round(drw.Clamps.Item(i).BasePointX, 3) & "," & Math.Round(drw.Clamps.Item(i).BasePointY, 3) & "," & state
    End If
    
Next i

End Sub

Public Sub readPositionsToArray()
Dim drw As Drawing
Set drw = App.ActiveDrawing
Dim temp() As String
Dim count As Integer
Dim size As Integer
Dim i As Integer
Dim j As Integer
Dim clmpCount As Integer
Dim posX As Integer
Dim posY As Integer
Dim posOp As Integer
Dim posState As Integer


If drw.Clamps.count = 30 Or drw.Clamps.count = 45 Then
    clmpCount = 15
ElseIf drw.Clamps.count = 32 Or drw.Clamps.count = 48 Then
    clmpCount = 16
Else
    MsgBox "Error"
    Exit Sub
End If


temp() = Split(drw.Clamps.Item(1).Attribute(data), ",")
count = UBound(temp)
size = ((count + 1) / 4) - 1

ReDim posFromData(size, clmpCount, 1, 1, 1)

For j = 1 To clmpCount
    posOp = 0
    posX = 1
    posY = 2
    posState = 3
    temp() = Split(drw.Clamps.Item(j).Attribute(data), ",")
    For i = 0 To size
        If IsNumeric(Right(temp(posOp), 2)) = True Then
            posFromData(i, j, 0, 0, 0) = Right(temp(posOp), 2)
        Else
            posFromData(i, j, 0, 0, 0) = Right(temp(posOp), 1)
        End If
        
        posFromData(i, j, 1, 0, 0) = temp(posX)
        posFromData(i, j, 0, 1, 0) = temp(posY)
        posFromData(i, j, 0, 0, 1) = temp(posState)
        
        posOp = posOp + 4
        posX = posX + 4
        posY = posY + 4
        posState = posState + 4
    Next i
Next j
                
End Sub

Public Sub positionClampsFromData()
Dim a As Integer
Dim b As Integer
Dim count As Integer
Dim j As Integer
Dim temp() As String
Dim size As Integer
Dim drw As Drawing
Set drw = App.ActiveDrawing

readPositionsToArray
currentPositions

If drw.Clamps.count = 30 Or drw.Clamps.count = 45 Then
    count = 15
ElseIf drw.Clamps.count = 32 Or drw.Clamps.count = 48 Then
    count = 16
Else
    MsgBox "Error"
    Exit Sub
End If


    For j = 1 To count
        drw.Clamps.Item(j).GeometryPath.MoveL posFromData(0, j, 1, 0, 0) - drw.Clamps.Item(j).BasePointX, posFromData(0, j, 0, 1, 0) - drw.Clamps.Item(j).BasePointY
        drw.Clamps.Item(j).GeometryPath.Visible = False
        drw.Clamps.Item(j).GeometryPath.Visible = True

        drw.Clamps.Item(j + count).GeometryPath.MoveL posFromData(0, j, 1, 0, 0) - drw.Clamps.Item(j + count).BasePointX, posFromData(0, j, 0, 1, 0) - drw.Clamps.Item(j + count).BasePointY
        drw.Clamps.Item(j + count).GeometryPath.Visible = False
        drw.Clamps.Item(j + count).GeometryPath.Visible = True
    Next j
    
    If drw.Clamps.count = 45 Then
        For j = 1 To count
            drw.Clamps.Item(j + count * 2).GeometryPath.MoveL posFromData(0, j, 1, 0, 0) - drw.Clamps.Item(j + count * 2).BasePointX, posFromData(0, j, 0, 1, 0) - drw.Clamps.Item(j + count * 2).BasePointY
            drw.Clamps.Item(j + count * 2).GeometryPath.Visible = False
            drw.Clamps.Item(j + count * 2).GeometryPath.Visible = True
        Next j
    End If
    If drw.Clamps.count = 48 Then
        For j = 1 To count
            drw.Clamps.Item(j + count * 2).GeometryPath.MoveL posFromData(0, j, 1, 0, 0) - drw.Clamps.Item(j + count * 2).BasePointX, posFromData(0, j, 0, 1, 0) - drw.Clamps.Item(j + count * 2).BasePointY
            drw.Clamps.Item(j + count * 2).GeometryPath.Visible = False
            drw.Clamps.Item(j + count * 2).GeometryPath.Visible = True
        Next j
    End If
    
temp() = Split(drw.Clamps.Item(1).Attribute(data), ",")
count = UBound(temp)
size = ((count + 1) / 4)

For b = 1 To drw.Operations.count
    drw.Operations.Item(b).Visible = False
Next b

If size = 1 Then
    For b = 1 To drw.Operations.count
        drw.Operations.Item(b).Visible = True
    Next b
    
Else

    For b = 1 To posFromData(1, 1, 0, 0, 0) - 1
        drw.Operations.Item(b).Visible = True
    Next b
    
End If

posFlag = 0

    frmPlayClamps.cmdPlayBackward.Enabled = False

If size > 1 Then
    frmPlayClamps.cmdPlayForward.Enabled = True
Else
    frmPlayClamps.cmdPlayForward.Enabled = False
End If

drw.Refresh
drw.RedrawShadedViews
drw.Redraw
drw.Refresh


frmPlayClamps.lblOpNum = "Op " & posFromData(0, 1, 0, 0, 0)

frmPlayClamps.Show

End Sub
Public Sub moveThroughPositions(direction As Integer)
Dim drw As Drawing
Set drw = App.ActiveDrawing
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim b As Integer
Dim temp() As String
Dim count As Integer
Dim size As Integer

temp() = Split(drw.Clamps.Item(1).Attribute(data), ",")
count = UBound(temp)
size = ((count + 1) / 4)

If direction = 1 Then
    i = posFlag + 1
    posFlag = i
End If
If direction = 0 Then
    i = posFlag - 1
    posFlag = i
End If

If posFlag = 0 Then
    frmPlayClamps.cmdPlayBackward.Enabled = False
Else
    frmPlayClamps.cmdPlayBackward.Enabled = True
End If
If posFlag = size - 1 Then
    frmPlayClamps.cmdPlayForward.Enabled = False
Else
    frmPlayClamps.cmdPlayForward.Enabled = True
End If


If drw.Clamps.count = 30 Or drw.Clamps.count = 45 Then
    count = 15
ElseIf drw.Clamps.count = 32 Or drw.Clamps.count = 48 Then
    count = 16
Else
    MsgBox "Error"
    Exit Sub
End If


    For j = 1 To count
        drw.Clamps.Item(j).GeometryPath.MoveL posFromData(i, j, 1, 0, 0) - drw.Clamps.Item(j).BasePointX, posFromData(i, j, 0, 1, 0) - drw.Clamps.Item(j).BasePointY
        drw.Clamps.Item(j).GeometryPath.Visible = False
        drw.Clamps.Item(j).GeometryPath.Visible = True

        drw.Clamps.Item(j + count).GeometryPath.MoveL posFromData(i, j, 1, 0, 0) - drw.Clamps.Item(j + count).BasePointX, posFromData(i, j, 0, 1, 0) - drw.Clamps.Item(j + count).BasePointY
        drw.Clamps.Item(j + count).GeometryPath.Visible = False
        drw.Clamps.Item(j + count).GeometryPath.Visible = True
    Next j
    
    If drw.Clamps.count = 45 Then
        For j = 1 To count
            drw.Clamps.Item(j + count * 2).GeometryPath.MoveL posFromData(i, j, 1, 0, 0) - drw.Clamps.Item(j + count * 2).BasePointX, posFromData(i, j, 0, 1, 0) - drw.Clamps.Item(j + count * 2).BasePointY
            drw.Clamps.Item(j + count * 2).GeometryPath.Visible = False
            drw.Clamps.Item(j + count * 2).GeometryPath.Visible = True
        Next j
    End If
    If drw.Clamps.count = 48 Then
        For j = 1 To count
            drw.Clamps.Item(j + count * 2).GeometryPath.MoveL posFromData(i, j, 1, 0, 0) - drw.Clamps.Item(j + count * 2).BasePointX, posFromData(i, j, 0, 1, 0) - drw.Clamps.Item(j + count * 2).BasePointY
            drw.Clamps.Item(j + count * 2).GeometryPath.Visible = False
            drw.Clamps.Item(j + count * 2).GeometryPath.Visible = True
        Next j
    End If

frmPlayClamps.lblOpNum = "Op " & posFromData(i, 1, 0, 0, 0)

For b = 1 To drw.Operations.count
    drw.Operations.Item(b).Visible = False
Next b

If size = 1 Then
    For b = 1 To drw.Operations.count
        drw.Operations.Item(b).Visible = True
    Next b
    
ElseIf i = size - 1 Then
    For b = posFromData(i, 1, 0, 0, 0) To drw.Operations.count
        drw.Operations.Item(b).Visible = True
    Next b
Else
    For b = posFromData(i, 1, 0, 0, 0) To posFromData(i + 1, 1, 0, 0, 0) - 1
        drw.Operations.Item(b).Visible = True
    Next b
    
End If

drw.Refresh
drw.RedrawShadedViews
drw.Redraw
drw.Refresh

End Sub

Public Sub removeClampOperation(opNum As Integer)
Dim drw As Drawing
Set drw = App.ActiveDrawing
Dim temp() As String
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim count As Integer
Dim clmpCount As Integer
Dim index As Integer
Dim size As Integer
Dim newTemp() As String

frmPlayClamps.Hide

temp() = Split(drw.Clamps.Item(1).Attribute(data), ",")
count = UBound(temp)
size = ((count + 1) / 4)

If drw.Clamps.count = 30 Or drw.Clamps.count = 45 Then
    clmpCount = 15
ElseIf drw.Clamps.count = 32 Or drw.Clamps.count = 48 Then
    clmpCount = 16
End If



For i = 1 To clmpCount

    index = 0
    temp() = Split(drw.Clamps.Item(i).Attribute(data), ",")
    
    For j = 0 To count
        
        If temp(j) = "OP" & opNum Then
            index = index + 1
            
            For k = j To j + 3
                temp(k) = ""
            Next k
        End If
        
    Next j
    
    If count - (index * 4) < 0 Then
        For j = 1 To clmpCount
            drw.Clamps.Item(j).Attribute(data) = ""
        Next j
        
        Exit Sub
    End If
       
    ReDim newTemp(count - (index * 4))
k = 0
    For j = 0 To count
        
        If temp(j) = "" Then
            k = k
        Else
            newTemp(k) = temp(j)
            k = k + 1
            
        End If
    Next j
    
    drw.Clamps.Item(i).Attribute(data) = newTemp(0)
    For j = 1 To count - (index * 4)
        drw.Clamps.Item(i).Attribute(data) = drw.Clamps.Item(i).Attribute(data) & "," & newTemp(j)
    Next j
    
Next i

positionClampsFromData

End Sub
Public Sub changeOperationNumber(opNum As Integer)
Dim drw As Drawing
Set drw = App.ActiveDrawing
Dim opNumNew As String
Dim temp() As String
Dim count As Integer
Dim size As Integer
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim clmpCount As Integer

frmPlayClamps.Hide

frmChangeOpNum.txtOpNum.Value = ""

frmChangeOpNum.Show

opNumNew = frmChangeOpNum.txtOpNum.Value

If opNumNew = "" Or opNumNew = "0" Or IsNumeric(opNumNew) = False Then
    Exit Sub
End If

temp() = Split(drw.Clamps.Item(1).Attribute(data), ",")
count = UBound(temp)
size = ((count + 1) / 4)

If drw.Clamps.count = 30 Or drw.Clamps.count = 45 Then
    clmpCount = 15
ElseIf drw.Clamps.count = 32 Or drw.Clamps.count = 48 Then
    clmpCount = 16
End If

For i = 1 To clmpCount
    temp() = Split(drw.Clamps.Item(i).Attribute(data), ",")
    For j = 0 To count
        If temp(j) = "OP" & opNum Then
            temp(j) = "OP" & opNumNew
        End If
    Next j
    drw.Clamps.Item(i).Attribute(data) = temp(0)
    For j = 1 To count
        drw.Clamps.Item(i).Attribute(data) = drw.Clamps.Item(i).Attribute(data) & "," & temp(j)
    Next j
    
Next i

positionClampsFromData
 
End Sub
Public Sub opsVisible()
Dim drw As Drawing
Set drw = App.ActiveDrawing
Dim i As Integer

For i = 1 To drw.Operations.count
    drw.Operations.Item(i).Visible = True
Next i

drw.Redraw
drw.Refresh

End Sub
Public Sub machineSelection()

frmMachine.Show

End Sub


Public Sub checkAndMove()
Dim drw As Drawing
Set drw = App.ActiveDrawing
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim a As Integer

    
If drw.Clamps.count = 30 Or drw.Clamps.count = 45 Then

    k = 1
    j = 0

    For i = 1 To 15
        If drw.Clamps.Item(i).BasePointX <> pos(j) Or drw.Clamps.Item(i).BasePointY <> pos(k) Then
                moveRowsAccord2 i
                checkBoundsAccord2
                checkAllignAccord2
                drw.Redraw
                drw.RedrawShadedViews
                currentPositions
            Exit For
        End If
    
        k = k + 2
        j = j + 2
    
    Next i

currentPositions
    
ElseIf drw.Clamps.count = 32 Or drw.Clamps.count = 48 Then

    k = 1
    j = 0

    For i = 1 To 16
        If drw.Clamps.Item(i).BasePointX <> pos(j) Or drw.Clamps.Item(i).BasePointY <> pos(k) Then
                moveRowsAccord1 i
                checkBoundsAccord1
                checkAllignAccord1
                drw.Redraw
                drw.RedrawShadedViews
                currentPositions
            Exit For
        End If
    
        k = k + 2
        j = j + 2
    
    Next i

currentPositions

ElseIf drw.Clamps.count = 18 Then

    k = 1
    j = 0

    For i = 1 To 9
        If drw.Clamps.Item(i).BasePointX <> pos(j) Or drw.Clamps.Item(i).BasePointY <> pos(k) Then
                moveRowsAccord25 i
                checkBoundsAccord25
                checkAllignAccord25
                drw.Redraw
                drw.RedrawShadedViews
                currentPositions
            Exit For
        End If
    
        k = k + 2
        j = j + 2
    
    Next i

currentPositions
End If


End Sub


Public Sub currentPositions()
Dim drw As Drawing
Set drw = App.ActiveDrawing
Dim i As Integer
Dim j As Integer

If drw.Clamps.count = 0 Then
    Exit Sub
ElseIf drw.Clamps.count = 30 Or drw.Clamps.count = 45 Then
    j = 0

    For i = 1 To 15
        pos(j) = drw.Clamps.Item(i).BasePointX
        j = j + 1
        pos(j) = drw.Clamps.Item(i).BasePointY
        j = j + 1
    Next i
    
ElseIf drw.Clamps.count = 32 Or drw.Clamps.count = 48 Then
    j = 0

    For i = 1 To 16
        pos(j) = drw.Clamps.Item(i).BasePointX
        j = j + 1
        pos(j) = drw.Clamps.Item(i).BasePointY
        j = j + 1
    Next i
    
ElseIf drw.Clamps.count = 18 Then
    j = 0

    For i = 1 To 9
        pos(j) = drw.Clamps.Item(i).BasePointX
        j = j + 1
        pos(j) = drw.Clamps.Item(i).BasePointY
        j = j + 1
    Next i
    
End If


End Sub





