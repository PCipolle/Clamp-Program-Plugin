Attribute VB_Name = "UserPositionClampsAccord25"
Public Sub moveRowsAccord25(i As Integer)

Dim drw As Drawing
Set drw = App.ActiveDrawing
Dim j As Integer
Dim k As Integer

j = i + 9

Dim Yspace As Double
Yspace = 6.25
Dim Xend As Double
Dim Yend As Double

Xend = drw.Clamps.Item(i).BasePointX
Yend = drw.Clamps.Item(i).BasePointY

If i = 1 Or i = 4 Or i = 7 Then
    drw.Clamps.Item(j).GeometryPath.MoveL Xend - drw.Clamps.Item(j).BasePointX, Yend - drw.Clamps.Item(j).BasePointY
    If drw.Clamps.Item(i + 1).BasePointY - Yend < Yspace Then
        drw.Clamps.Item(i + 1).GeometryPath.MoveL Xend - drw.Clamps.Item(i + 1).BasePointX, Yend + Yspace - drw.Clamps.Item(i + 1).BasePointY
        drw.Clamps.Item(j + 1).GeometryPath.MoveL Xend - drw.Clamps.Item(j + 1).BasePointX, Yend + Yspace - drw.Clamps.Item(j + 1).BasePointY
    Else
        drw.Clamps.Item(i + 1).GeometryPath.MoveL Xend - drw.Clamps.Item(i + 1).BasePointX, 0
        drw.Clamps.Item(j + 1).GeometryPath.MoveL Xend - drw.Clamps.Item(j + 1).BasePointX, 0
    End If
    If drw.Clamps.Item(i + 2).BasePointY - drw.Clamps.Item(i + 1).BasePointY < Yspace Then
        drw.Clamps.Item(i + 2).GeometryPath.MoveL Xend - drw.Clamps.Item(i + 2).BasePointX, drw.Clamps.Item(i + 1).BasePointY + Yspace - drw.Clamps.Item(i + 2).BasePointY
        drw.Clamps.Item(j + 2).GeometryPath.MoveL Xend - drw.Clamps.Item(j + 2).BasePointX, drw.Clamps.Item(j + 1).BasePointY + Yspace - drw.Clamps.Item(j + 2).BasePointY
    Else
        drw.Clamps.Item(i + 2).GeometryPath.MoveL Xend - drw.Clamps.Item(i + 2).BasePointX, 0
        drw.Clamps.Item(j + 2).GeometryPath.MoveL Xend - drw.Clamps.Item(j + 2).BasePointX, 0
    End If
End If
If i = 2 Or i = 5 Or i = 8 Then
    drw.Clamps.Item(j).GeometryPath.MoveL Xend - drw.Clamps.Item(j).BasePointX, Yend - drw.Clamps.Item(j).BasePointY
    If drw.Clamps.Item(i + 1).BasePointY - Yend < Yspace Then
        drw.Clamps.Item(i + 1).GeometryPath.MoveL Xend - drw.Clamps.Item(i + 1).BasePointX, Yend + Yspace - drw.Clamps.Item(i + 1).BasePointY
        drw.Clamps.Item(j + 1).GeometryPath.MoveL Xend - drw.Clamps.Item(j + 1).BasePointX, Yend + Yspace - drw.Clamps.Item(j + 1).BasePointY
    Else
        drw.Clamps.Item(i + 1).GeometryPath.MoveL Xend - drw.Clamps.Item(i + 1).BasePointX, 0
        drw.Clamps.Item(j + 1).GeometryPath.MoveL Xend - drw.Clamps.Item(j + 1).BasePointX, 0
    End If
    If Yend - drw.Clamps.Item(i - 1).BasePointY < Yspace Then
        drw.Clamps.Item(i - 1).GeometryPath.MoveL Xend - drw.Clamps.Item(i - 1).BasePointX, Yend - Yspace - drw.Clamps.Item(i - 1).BasePointY
        drw.Clamps.Item(j - 1).GeometryPath.MoveL Xend - drw.Clamps.Item(j - 1).BasePointX, Yend - Yspace - drw.Clamps.Item(j - 1).BasePointY
    Else
        drw.Clamps.Item(i - 1).GeometryPath.MoveL Xend - drw.Clamps.Item(i - 1).BasePointX, 0
        drw.Clamps.Item(j - 1).GeometryPath.MoveL Xend - drw.Clamps.Item(j - 1).BasePointX, 0
    End If
End If
If i = 3 Or i = 6 Or i = 9 Then
    drw.Clamps.Item(j).GeometryPath.MoveL Xend - drw.Clamps.Item(j).BasePointX, Yend - drw.Clamps.Item(j).BasePointY
    If Yend - drw.Clamps.Item(i - 1).BasePointY < Yspace Then
        drw.Clamps.Item(i - 1).GeometryPath.MoveL Xend - drw.Clamps.Item(i - 1).BasePointX, Yend - Yspace - drw.Clamps.Item(i - 1).BasePointY
        drw.Clamps.Item(j - 1).GeometryPath.MoveL Xend - drw.Clamps.Item(j - 1).BasePointX, Yend - Yspace - drw.Clamps.Item(j - 1).BasePointY
    Else
        drw.Clamps.Item(i - 1).GeometryPath.MoveL Xend - drw.Clamps.Item(i - 1).BasePointX, 0
        drw.Clamps.Item(j - 1).GeometryPath.MoveL Xend - drw.Clamps.Item(j - 1).BasePointX, 0
    End If
    If drw.Clamps.Item(i - 1).BasePointY - drw.Clamps.Item(i - 2).BasePointY < Yspace Then
        drw.Clamps.Item(i - 2).GeometryPath.MoveL Xend - drw.Clamps.Item(i - 2).BasePointX, drw.Clamps.Item(i - 1).BasePointY - Yspace - drw.Clamps.Item(i - 2).BasePointY
        drw.Clamps.Item(j - 2).GeometryPath.MoveL Xend - drw.Clamps.Item(j - 2).BasePointX, drw.Clamps.Item(i - 1).BasePointY - Yspace - drw.Clamps.Item(j - 2).BasePointY
    Else
        drw.Clamps.Item(i - 2).GeometryPath.MoveL Xend - drw.Clamps.Item(i - 2).BasePointX, 0
        drw.Clamps.Item(j - 2).GeometryPath.MoveL Xend - drw.Clamps.Item(j - 2).BasePointX, 0
    End If
End If
    
moveBarsAccord25 i

End Sub
Public Sub moveBarsAccord25(i As Integer)

If i = 1 Or i = 2 Or i = 3 Then
    bar1Accord25Move i
End If
If i = 4 Or i = 5 Or i = 6 Then
    bar2Accord25Move i
End If
If i = 7 Or i = 8 Or i = 9 Then
    bar3Accord25Move i
End If

End Sub
Public Sub bar1Accord25Move(i As Integer)
Dim drw As Drawing
Set drw = App.ActiveDrawing

Dim Xspace As Double
Xspace = 6.5
Dim Xend As Double

Xend = drw.Clamps.Item(i).BasePointX

If drw.Clamps.Item(4).BasePointX - Xend < Xspace Then
    drw.Clamps.Item(4).GeometryPath.MoveL Xend + Xspace - drw.Clamps.Item(15).BasePointX, 0
    drw.Clamps.Item(5).GeometryPath.MoveL Xend + Xspace - drw.Clamps.Item(15).BasePointX, 0
    drw.Clamps.Item(6).GeometryPath.MoveL Xend + Xspace - drw.Clamps.Item(15).BasePointX, 0
    drw.Clamps.Item(13).GeometryPath.MoveL Xend + Xspace - drw.Clamps.Item(15).BasePointX, 0
    drw.Clamps.Item(14).GeometryPath.MoveL Xend + Xspace - drw.Clamps.Item(15).BasePointX, 0
    drw.Clamps.Item(15).GeometryPath.MoveL Xend + Xspace - drw.Clamps.Item(15).BasePointX, 0
End If
If drw.Clamps.Item(7).BasePointX - drw.Clamps.Item(4).BasePointX < Xspace Then
    drw.Clamps.Item(7).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX + Xspace - drw.Clamps.Item(18).BasePointX, 0
    drw.Clamps.Item(8).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX + Xspace - drw.Clamps.Item(18).BasePointX, 0
    drw.Clamps.Item(9).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX + Xspace - drw.Clamps.Item(18).BasePointX, 0
    drw.Clamps.Item(16).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX + Xspace - drw.Clamps.Item(18).BasePointX, 0
    drw.Clamps.Item(17).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX + Xspace - drw.Clamps.Item(18).BasePointX, 0
    drw.Clamps.Item(18).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX + Xspace - drw.Clamps.Item(18).BasePointX, 0
End If

End Sub
Public Sub bar2Accord25Move(i As Integer)
Dim drw As Drawing
Set drw = App.ActiveDrawing

Dim Xspace As Double
Xspace = 6.5
Dim Xend As Double

Xend = drw.Clamps.Item(i).BasePointX

If drw.Clamps.Item(7).BasePointX - drw.Clamps.Item(4).BasePointX < Xspace Then
    drw.Clamps.Item(7).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX + Xspace - drw.Clamps.Item(18).BasePointX, 0
    drw.Clamps.Item(8).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX + Xspace - drw.Clamps.Item(18).BasePointX, 0
    drw.Clamps.Item(9).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX + Xspace - drw.Clamps.Item(18).BasePointX, 0
    drw.Clamps.Item(16).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX + Xspace - drw.Clamps.Item(18).BasePointX, 0
    drw.Clamps.Item(17).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX + Xspace - drw.Clamps.Item(18).BasePointX, 0
    drw.Clamps.Item(18).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX + Xspace - drw.Clamps.Item(18).BasePointX, 0
End If
If Xend - drw.Clamps.Item(1).BasePointX < Xspace Then
    drw.Clamps.Item(1).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(12).BasePointX, 0
    drw.Clamps.Item(2).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(12).BasePointX, 0
    drw.Clamps.Item(3).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(12).BasePointX, 0
    drw.Clamps.Item(10).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(12).BasePointX, 0
    drw.Clamps.Item(11).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(12).BasePointX, 0
    drw.Clamps.Item(12).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(12).BasePointX, 0
End If


End Sub
Public Sub bar3Accord25Move(i As Integer)
Dim drw As Drawing
Set drw = App.ActiveDrawing

Dim Xspace As Double
Xspace = 6.5
Dim Xend As Double

Xend = drw.Clamps.Item(i).BasePointX

If Xend - drw.Clamps.Item(4).BasePointX < Xspace Then
    drw.Clamps.Item(4).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(15).BasePointX, 0
    drw.Clamps.Item(5).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(15).BasePointX, 0
    drw.Clamps.Item(6).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(15).BasePointX, 0
    drw.Clamps.Item(13).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(15).BasePointX, 0
    drw.Clamps.Item(14).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(15).BasePointX, 0
    drw.Clamps.Item(15).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(15).BasePointX, 0
End If
If drw.Clamps.Item(4).BasePointX - drw.Clamps.Item(1).BasePointX < Xspace Then
    drw.Clamps.Item(1).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - Xspace - drw.Clamps.Item(12).BasePointX, 0
    drw.Clamps.Item(2).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - Xspace - drw.Clamps.Item(12).BasePointX, 0
    drw.Clamps.Item(3).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - Xspace - drw.Clamps.Item(12).BasePointX, 0
    drw.Clamps.Item(10).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - Xspace - drw.Clamps.Item(12).BasePointX, 0
    drw.Clamps.Item(11).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - Xspace - drw.Clamps.Item(12).BasePointX, 0
    drw.Clamps.Item(12).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - Xspace - drw.Clamps.Item(12).BasePointX, 0
End If


End Sub



Public Sub checkBoundsAccord25()
Dim drw As Drawing
Set drw = App.ActiveDrawing

Dim Xmin As Double
Dim Ymin As Double
Dim Xmax As Double
Dim Ymax As Double

Xmin = 3.5
Xmax = 69
Ymin = 3.2
Ymax = 49

If drw.Clamps.Item(1).BasePointY < Ymin Or drw.Clamps.Item(4).BasePointY < Ymin Or drw.Clamps.Item(7).BasePointY < Ymin Then
    MsgBox "Row 1 out of bounds! Minimum is 3.2 inches"
    drw.Undo
End If
If drw.Clamps.Item(3).BasePointY > Ymax Or drw.Clamps.Item(6).BasePointY > Ymax Or drw.Clamps.Item(9).BasePointY > Ymax Then
    MsgBox "Row 3 out of bounds! Maximum is 49 inches"
    drw.Undo
End If
If drw.Clamps.Item(1).BasePointX < Xmin Then
    MsgBox "Bar 1 out of bounds! Minimum is 3.5 inches"
    drw.Undo
End If
If drw.Clamps.Item(7).BasePointX > Xmax Then
    MsgBox "Bar 3 out of bounds! Minimum is 69 inches"
    drw.Undo
End If

End Sub
Public Sub checkAllignAccord25()
Dim drw As Drawing
Set drw = App.ActiveDrawing

If drw.Clamps.Item(1).BasePointX <> drw.Clamps.Item(2).BasePointX Or drw.Clamps.Item(2).BasePointX <> drw.Clamps.Item(3).BasePointX Then
    drw.Clamps.Item(2).GeometryPath.MoveL drw.Clamps.Item(1).BasePointX - drw.Clamps.Item(2).BasePointX, 0
    drw.Clamps.Item(3).GeometryPath.MoveL drw.Clamps.Item(1).BasePointX - drw.Clamps.Item(3).BasePointX, 0
    drw.Clamps.Item(11).GeometryPath.MoveL drw.Clamps.Item(1).BasePointX - drw.Clamps.Item(11).BasePointX, 0
    drw.Clamps.Item(12).GeometryPath.MoveL drw.Clamps.Item(1).BasePointX - drw.Clamps.Item(12).BasePointX, 0
End If
If drw.Clamps.Item(4).BasePointX <> drw.Clamps.Item(5).BasePointX Or drw.Clamps.Item(5).BasePointX <> drw.Clamps.Item(6).BasePointX Then
    drw.Clamps.Item(5).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - drw.Clamps.Item(5).BasePointX, 0
    drw.Clamps.Item(6).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - drw.Clamps.Item(6).BasePointX, 0
    drw.Clamps.Item(14).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - drw.Clamps.Item(14).BasePointX, 0
    drw.Clamps.Item(15).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - drw.Clamps.Item(15).BasePointX, 0
End If
If drw.Clamps.Item(7).BasePointX <> drw.Clamps.Item(8).BasePointX Or drw.Clamps.Item(8).BasePointX <> drw.Clamps.Item(9).BasePointX Then
    drw.Clamps.Item(8).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX - drw.Clamps.Item(8).BasePointX, 0
    drw.Clamps.Item(9).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX - drw.Clamps.Item(9).BasePointX, 0
    drw.Clamps.Item(17).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX - drw.Clamps.Item(17).BasePointX, 0
    drw.Clamps.Item(18).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX - drw.Clamps.Item(18).BasePointX, 0
End If

End Sub
