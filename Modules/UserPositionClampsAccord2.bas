Attribute VB_Name = "UserPositionClampsAccord2"
Public Sub moveRowsAccord2(i As Integer)

Dim drw As Drawing
Set drw = App.ActiveDrawing
Dim j As Integer
Dim k As Integer

    j = i + 15
    k = i + 30

Dim Yspace As Double
Yspace = 6.5
Dim Xend As Double
Dim Yend As Double

Xend = drw.Clamps.Item(i).BasePointX
Yend = drw.Clamps.Item(i).BasePointY

If drw.Clamps.count = 30 Then

    If i = 1 Or i = 4 Or i = 7 Or i = 10 Or i = 13 Then
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
    If i = 2 Or i = 5 Or i = 8 Or i = 11 Or i = 14 Then
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
    If i = 3 Or i = 6 Or i = 9 Or i = 12 Or i = 15 Then
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
    
ElseIf drw.Clamps.count = 45 Then


    If i = 1 Or i = 4 Or i = 7 Or i = 10 Or i = 13 Then
        drw.Clamps.Item(j).GeometryPath.MoveL Xend - drw.Clamps.Item(j).BasePointX, Yend - drw.Clamps.Item(j).BasePointY
        drw.Clamps.Item(k).GeometryPath.MoveL Xend - drw.Clamps.Item(k).BasePointX, Yend - drw.Clamps.Item(k).BasePointY
        If drw.Clamps.Item(i + 1).BasePointY - Yend < Yspace Then
            drw.Clamps.Item(i + 1).GeometryPath.MoveL Xend - drw.Clamps.Item(i + 1).BasePointX, Yend + Yspace - drw.Clamps.Item(i + 1).BasePointY
            drw.Clamps.Item(j + 1).GeometryPath.MoveL Xend - drw.Clamps.Item(j + 1).BasePointX, Yend + Yspace - drw.Clamps.Item(j + 1).BasePointY
            drw.Clamps.Item(k + 1).GeometryPath.MoveL Xend - drw.Clamps.Item(k + 1).BasePointX, Yend + Yspace - drw.Clamps.Item(k + 1).BasePointY
        Else
            drw.Clamps.Item(i + 1).GeometryPath.MoveL Xend - drw.Clamps.Item(i + 1).BasePointX, 0
            drw.Clamps.Item(j + 1).GeometryPath.MoveL Xend - drw.Clamps.Item(j + 1).BasePointX, 0
            drw.Clamps.Item(k + 1).GeometryPath.MoveL Xend - drw.Clamps.Item(k + 1).BasePointX, 0
        End If
        If drw.Clamps.Item(i + 2).BasePointY - drw.Clamps.Item(i + 1).BasePointY < Yspace Then
            drw.Clamps.Item(i + 2).GeometryPath.MoveL Xend - drw.Clamps.Item(i + 2).BasePointX, drw.Clamps.Item(i + 1).BasePointY + Yspace - drw.Clamps.Item(i + 2).BasePointY
            drw.Clamps.Item(j + 2).GeometryPath.MoveL Xend - drw.Clamps.Item(j + 2).BasePointX, drw.Clamps.Item(j + 1).BasePointY + Yspace - drw.Clamps.Item(j + 2).BasePointY
            drw.Clamps.Item(k + 2).GeometryPath.MoveL Xend - drw.Clamps.Item(k + 2).BasePointX, drw.Clamps.Item(k + 1).BasePointY + Yspace - drw.Clamps.Item(k + 2).BasePointY
        Else
            drw.Clamps.Item(i + 2).GeometryPath.MoveL Xend - drw.Clamps.Item(i + 2).BasePointX, 0
            drw.Clamps.Item(j + 2).GeometryPath.MoveL Xend - drw.Clamps.Item(j + 2).BasePointX, 0
            drw.Clamps.Item(k + 2).GeometryPath.MoveL Xend - drw.Clamps.Item(k + 2).BasePointX, 0
        End If
    End If
    If i = 2 Or i = 5 Or i = 8 Or i = 11 Or i = 14 Then
        drw.Clamps.Item(j).GeometryPath.MoveL Xend - drw.Clamps.Item(j).BasePointX, Yend - drw.Clamps.Item(j).BasePointY
        drw.Clamps.Item(k).GeometryPath.MoveL Xend - drw.Clamps.Item(k).BasePointX, Yend - drw.Clamps.Item(k).BasePointY
        If drw.Clamps.Item(i + 1).BasePointY - Yend < Yspace Then
            drw.Clamps.Item(i + 1).GeometryPath.MoveL Xend - drw.Clamps.Item(i + 1).BasePointX, Yend + Yspace - drw.Clamps.Item(i + 1).BasePointY
            drw.Clamps.Item(j + 1).GeometryPath.MoveL Xend - drw.Clamps.Item(j + 1).BasePointX, Yend + Yspace - drw.Clamps.Item(j + 1).BasePointY
            drw.Clamps.Item(k + 1).GeometryPath.MoveL Xend - drw.Clamps.Item(k + 1).BasePointX, Yend + Yspace - drw.Clamps.Item(k + 1).BasePointY
        Else
            drw.Clamps.Item(i + 1).GeometryPath.MoveL Xend - drw.Clamps.Item(i + 1).BasePointX, 0
            drw.Clamps.Item(j + 1).GeometryPath.MoveL Xend - drw.Clamps.Item(j + 1).BasePointX, 0
            drw.Clamps.Item(k + 1).GeometryPath.MoveL Xend - drw.Clamps.Item(k + 1).BasePointX, 0
        End If
        If Yend - drw.Clamps.Item(i - 1).BasePointY < Yspace Then
            drw.Clamps.Item(i - 1).GeometryPath.MoveL Xend - drw.Clamps.Item(i - 1).BasePointX, Yend - Yspace - drw.Clamps.Item(i - 1).BasePointY
            drw.Clamps.Item(j - 1).GeometryPath.MoveL Xend - drw.Clamps.Item(j - 1).BasePointX, Yend - Yspace - drw.Clamps.Item(j - 1).BasePointY
            drw.Clamps.Item(k - 1).GeometryPath.MoveL Xend - drw.Clamps.Item(k - 1).BasePointX, Yend - Yspace - drw.Clamps.Item(k - 1).BasePointY
        Else
            drw.Clamps.Item(i - 1).GeometryPath.MoveL Xend - drw.Clamps.Item(i - 1).BasePointX, 0
            drw.Clamps.Item(j - 1).GeometryPath.MoveL Xend - drw.Clamps.Item(j - 1).BasePointX, 0
            drw.Clamps.Item(k - 1).GeometryPath.MoveL Xend - drw.Clamps.Item(k - 1).BasePointX, 0
        End If
    End If
    If i = 3 Or i = 6 Or i = 9 Or i = 12 Or i = 15 Then
        drw.Clamps.Item(j).GeometryPath.MoveL Xend - drw.Clamps.Item(j).BasePointX, Yend - drw.Clamps.Item(j).BasePointY
        drw.Clamps.Item(k).GeometryPath.MoveL Xend - drw.Clamps.Item(k).BasePointX, Yend - drw.Clamps.Item(k).BasePointY
        If Yend - drw.Clamps.Item(i - 1).BasePointY < Yspace Then
            drw.Clamps.Item(i - 1).GeometryPath.MoveL Xend - drw.Clamps.Item(i - 1).BasePointX, Yend - Yspace - drw.Clamps.Item(i - 1).BasePointY
            drw.Clamps.Item(j - 1).GeometryPath.MoveL Xend - drw.Clamps.Item(j - 1).BasePointX, Yend - Yspace - drw.Clamps.Item(j - 1).BasePointY
            drw.Clamps.Item(k - 1).GeometryPath.MoveL Xend - drw.Clamps.Item(k - 1).BasePointX, Yend - Yspace - drw.Clamps.Item(k - 1).BasePointY
        Else
            drw.Clamps.Item(i - 1).GeometryPath.MoveL Xend - drw.Clamps.Item(i - 1).BasePointX, 0
            drw.Clamps.Item(j - 1).GeometryPath.MoveL Xend - drw.Clamps.Item(j - 1).BasePointX, 0
            drw.Clamps.Item(k - 1).GeometryPath.MoveL Xend - drw.Clamps.Item(k - 1).BasePointX, 0
        End If
        If drw.Clamps.Item(i - 1).BasePointY - drw.Clamps.Item(i - 2).BasePointY < Yspace Then
            drw.Clamps.Item(i - 2).GeometryPath.MoveL Xend - drw.Clamps.Item(i - 2).BasePointX, drw.Clamps.Item(i - 1).BasePointY - Yspace - drw.Clamps.Item(i - 2).BasePointY
            drw.Clamps.Item(j - 2).GeometryPath.MoveL Xend - drw.Clamps.Item(j - 2).BasePointX, drw.Clamps.Item(i - 1).BasePointY - Yspace - drw.Clamps.Item(j - 2).BasePointY
            drw.Clamps.Item(k - 2).GeometryPath.MoveL Xend - drw.Clamps.Item(k - 2).BasePointX, drw.Clamps.Item(i - 1).BasePointY - Yspace - drw.Clamps.Item(k - 2).BasePointY
        Else
            drw.Clamps.Item(i - 2).GeometryPath.MoveL Xend - drw.Clamps.Item(i - 2).BasePointX, 0
            drw.Clamps.Item(j - 2).GeometryPath.MoveL Xend - drw.Clamps.Item(j - 2).BasePointX, 0
            drw.Clamps.Item(k - 2).GeometryPath.MoveL Xend - drw.Clamps.Item(k - 2).BasePointX, 0
        End If
    End If

End If

moveBarsAccord2 i

End Sub
Public Sub moveBarsAccord2(i As Integer)

If i = 1 Or i = 2 Or i = 3 Then
    bar1Accord2Move i
End If
If i = 4 Or i = 5 Or i = 6 Then
    bar2Accord2Move i
End If
If i = 7 Or i = 8 Or i = 9 Then
    bar3Accord2Move i
End If
If i = 10 Or i = 11 Or i = 12 Then
    bar4Accord2Move i
End If
If i = 13 Or i = 14 Or i = 15 Then
    bar5Accord2Move i
End If

End Sub
Public Sub bar1Accord2Move(i As Integer)
Dim drw As Drawing
Set drw = App.ActiveDrawing

Dim Xspace As Double
Xspace = 7.125
Dim Xend As Double


Xend = drw.Clamps.Item(i).BasePointX
If drw.Clamps.count = 30 Then

    If drw.Clamps.Item(4).BasePointX - Xend < Xspace Then
        drw.Clamps.Item(4).GeometryPath.MoveL Xend + Xspace - drw.Clamps.Item(21).BasePointX, 0
        drw.Clamps.Item(5).GeometryPath.MoveL Xend + Xspace - drw.Clamps.Item(21).BasePointX, 0
        drw.Clamps.Item(6).GeometryPath.MoveL Xend + Xspace - drw.Clamps.Item(21).BasePointX, 0
        drw.Clamps.Item(19).GeometryPath.MoveL Xend + Xspace - drw.Clamps.Item(21).BasePointX, 0
        drw.Clamps.Item(20).GeometryPath.MoveL Xend + Xspace - drw.Clamps.Item(21).BasePointX, 0
        drw.Clamps.Item(21).GeometryPath.MoveL Xend + Xspace - drw.Clamps.Item(21).BasePointX, 0
    End If
    If drw.Clamps.Item(7).BasePointX - drw.Clamps.Item(4).BasePointX < Xspace Then
        drw.Clamps.Item(7).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX + Xspace - drw.Clamps.Item(24).BasePointX, 0
        drw.Clamps.Item(8).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX + Xspace - drw.Clamps.Item(24).BasePointX, 0
        drw.Clamps.Item(9).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX + Xspace - drw.Clamps.Item(24).BasePointX, 0
        drw.Clamps.Item(22).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX + Xspace - drw.Clamps.Item(24).BasePointX, 0
        drw.Clamps.Item(23).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX + Xspace - drw.Clamps.Item(24).BasePointX, 0
        drw.Clamps.Item(24).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX + Xspace - drw.Clamps.Item(24).BasePointX, 0
    End If
    If drw.Clamps.Item(10).BasePointX - drw.Clamps.Item(7).BasePointX < Xspace Then
        drw.Clamps.Item(10).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX + Xspace - drw.Clamps.Item(27).BasePointX, 0
        drw.Clamps.Item(11).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX + Xspace - drw.Clamps.Item(27).BasePointX, 0
        drw.Clamps.Item(12).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX + Xspace - drw.Clamps.Item(27).BasePointX, 0
        drw.Clamps.Item(25).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX + Xspace - drw.Clamps.Item(27).BasePointX, 0
        drw.Clamps.Item(26).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX + Xspace - drw.Clamps.Item(27).BasePointX, 0
        drw.Clamps.Item(27).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX + Xspace - drw.Clamps.Item(27).BasePointX, 0
    End If
    If drw.Clamps.Item(13).BasePointX - drw.Clamps.Item(10).BasePointX < Xspace Then
        drw.Clamps.Item(13).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(30).BasePointX, 0
        drw.Clamps.Item(14).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(30).BasePointX, 0
        drw.Clamps.Item(15).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(30).BasePointX, 0
        drw.Clamps.Item(28).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(30).BasePointX, 0
        drw.Clamps.Item(29).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(30).BasePointX, 0
        drw.Clamps.Item(30).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(30).BasePointX, 0
    End If
ElseIf drw.Clamps.count = 45 Then

    If drw.Clamps.Item(4).BasePointX - Xend < Xspace Then
        drw.Clamps.Item(4).GeometryPath.MoveL Xend + Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(5).GeometryPath.MoveL Xend + Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(6).GeometryPath.MoveL Xend + Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(19).GeometryPath.MoveL Xend + Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(20).GeometryPath.MoveL Xend + Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(21).GeometryPath.MoveL Xend + Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(34).GeometryPath.MoveL Xend + Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(35).GeometryPath.MoveL Xend + Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(36).GeometryPath.MoveL Xend + Xspace - drw.Clamps.Item(36).BasePointX, 0
    End If
    If drw.Clamps.Item(7).BasePointX - drw.Clamps.Item(4).BasePointX < Xspace Then
        drw.Clamps.Item(7).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX + Xspace - drw.Clamps.Item(39).BasePointX, 0
        drw.Clamps.Item(8).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX + Xspace - drw.Clamps.Item(39).BasePointX, 0
        drw.Clamps.Item(9).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX + Xspace - drw.Clamps.Item(39).BasePointX, 0
        drw.Clamps.Item(22).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX + Xspace - drw.Clamps.Item(39).BasePointX, 0
        drw.Clamps.Item(23).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX + Xspace - drw.Clamps.Item(39).BasePointX, 0
        drw.Clamps.Item(24).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX + Xspace - drw.Clamps.Item(39).BasePointX, 0
        drw.Clamps.Item(37).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX + Xspace - drw.Clamps.Item(39).BasePointX, 0
        drw.Clamps.Item(38).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX + Xspace - drw.Clamps.Item(39).BasePointX, 0
        drw.Clamps.Item(39).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX + Xspace - drw.Clamps.Item(39).BasePointX, 0
    End If
    If drw.Clamps.Item(10).BasePointX - drw.Clamps.Item(7).BasePointX < Xspace Then
        drw.Clamps.Item(10).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX + Xspace - drw.Clamps.Item(42).BasePointX, 0
        drw.Clamps.Item(11).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX + Xspace - drw.Clamps.Item(42).BasePointX, 0
        drw.Clamps.Item(12).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX + Xspace - drw.Clamps.Item(42).BasePointX, 0
        drw.Clamps.Item(25).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX + Xspace - drw.Clamps.Item(42).BasePointX, 0
        drw.Clamps.Item(26).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX + Xspace - drw.Clamps.Item(42).BasePointX, 0
        drw.Clamps.Item(27).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX + Xspace - drw.Clamps.Item(42).BasePointX, 0
        drw.Clamps.Item(40).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX + Xspace - drw.Clamps.Item(42).BasePointX, 0
        drw.Clamps.Item(41).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX + Xspace - drw.Clamps.Item(42).BasePointX, 0
        drw.Clamps.Item(42).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX + Xspace - drw.Clamps.Item(42).BasePointX, 0
    End If
    If drw.Clamps.Item(13).BasePointX - drw.Clamps.Item(10).BasePointX < Xspace Then
        drw.Clamps.Item(13).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(45).BasePointX, 0
        drw.Clamps.Item(14).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(45).BasePointX, 0
        drw.Clamps.Item(15).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(45).BasePointX, 0
        drw.Clamps.Item(28).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(45).BasePointX, 0
        drw.Clamps.Item(29).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(45).BasePointX, 0
        drw.Clamps.Item(30).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(45).BasePointX, 0
        drw.Clamps.Item(43).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(45).BasePointX, 0
        drw.Clamps.Item(44).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(45).BasePointX, 0
        drw.Clamps.Item(45).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(45).BasePointX, 0
    End If
End If




End Sub
Public Sub bar2Accord2Move(i As Integer)
Dim drw As Drawing
Set drw = App.ActiveDrawing

Dim Xspace As Double
Xspace = 7.125
Dim Xend As Double


Xend = drw.Clamps.Item(i).BasePointX
If drw.Clamps.count = 30 Then

    If drw.Clamps.Item(7).BasePointX - drw.Clamps.Item(4).BasePointX < Xspace Then
        drw.Clamps.Item(7).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX + Xspace - drw.Clamps.Item(24).BasePointX, 0
        drw.Clamps.Item(8).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX + Xspace - drw.Clamps.Item(24).BasePointX, 0
        drw.Clamps.Item(9).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX + Xspace - drw.Clamps.Item(24).BasePointX, 0
        drw.Clamps.Item(22).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX + Xspace - drw.Clamps.Item(24).BasePointX, 0
        drw.Clamps.Item(23).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX + Xspace - drw.Clamps.Item(24).BasePointX, 0
        drw.Clamps.Item(24).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX + Xspace - drw.Clamps.Item(24).BasePointX, 0
    End If
    If drw.Clamps.Item(10).BasePointX - drw.Clamps.Item(7).BasePointX < Xspace Then
        drw.Clamps.Item(10).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX + Xspace - drw.Clamps.Item(27).BasePointX, 0
        drw.Clamps.Item(11).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX + Xspace - drw.Clamps.Item(27).BasePointX, 0
        drw.Clamps.Item(12).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX + Xspace - drw.Clamps.Item(27).BasePointX, 0
        drw.Clamps.Item(25).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX + Xspace - drw.Clamps.Item(27).BasePointX, 0
        drw.Clamps.Item(26).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX + Xspace - drw.Clamps.Item(27).BasePointX, 0
        drw.Clamps.Item(27).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX + Xspace - drw.Clamps.Item(27).BasePointX, 0
    End If
    If drw.Clamps.Item(13).BasePointX - drw.Clamps.Item(10).BasePointX < Xspace Then
        drw.Clamps.Item(13).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(30).BasePointX, 0
        drw.Clamps.Item(14).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(30).BasePointX, 0
        drw.Clamps.Item(15).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(30).BasePointX, 0
        drw.Clamps.Item(28).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(30).BasePointX, 0
        drw.Clamps.Item(29).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(30).BasePointX, 0
        drw.Clamps.Item(30).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(30).BasePointX, 0
    End If
    If Xend - drw.Clamps.Item(1).BasePointX < Xspace Then
        drw.Clamps.Item(1).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(18).BasePointX, 0
        drw.Clamps.Item(2).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(18).BasePointX, 0
        drw.Clamps.Item(3).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(18).BasePointX, 0
        drw.Clamps.Item(16).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(18).BasePointX, 0
        drw.Clamps.Item(17).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(18).BasePointX, 0
        drw.Clamps.Item(18).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(18).BasePointX, 0
    End If
ElseIf drw.Clamps.count = 45 Then

    If drw.Clamps.Item(7).BasePointX - drw.Clamps.Item(4).BasePointX < Xspace Then
        drw.Clamps.Item(7).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX + Xspace - drw.Clamps.Item(39).BasePointX, 0
        drw.Clamps.Item(8).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX + Xspace - drw.Clamps.Item(39).BasePointX, 0
        drw.Clamps.Item(9).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX + Xspace - drw.Clamps.Item(39).BasePointX, 0
        drw.Clamps.Item(22).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX + Xspace - drw.Clamps.Item(39).BasePointX, 0
        drw.Clamps.Item(23).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX + Xspace - drw.Clamps.Item(39).BasePointX, 0
        drw.Clamps.Item(24).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX + Xspace - drw.Clamps.Item(39).BasePointX, 0
        drw.Clamps.Item(37).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX + Xspace - drw.Clamps.Item(39).BasePointX, 0
        drw.Clamps.Item(38).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX + Xspace - drw.Clamps.Item(39).BasePointX, 0
        drw.Clamps.Item(39).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX + Xspace - drw.Clamps.Item(39).BasePointX, 0
    End If
    If drw.Clamps.Item(10).BasePointX - drw.Clamps.Item(7).BasePointX < Xspace Then
        drw.Clamps.Item(10).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX + Xspace - drw.Clamps.Item(42).BasePointX, 0
        drw.Clamps.Item(11).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX + Xspace - drw.Clamps.Item(42).BasePointX, 0
        drw.Clamps.Item(12).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX + Xspace - drw.Clamps.Item(42).BasePointX, 0
        drw.Clamps.Item(25).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX + Xspace - drw.Clamps.Item(42).BasePointX, 0
        drw.Clamps.Item(26).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX + Xspace - drw.Clamps.Item(42).BasePointX, 0
        drw.Clamps.Item(27).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX + Xspace - drw.Clamps.Item(42).BasePointX, 0
        drw.Clamps.Item(40).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX + Xspace - drw.Clamps.Item(42).BasePointX, 0
        drw.Clamps.Item(41).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX + Xspace - drw.Clamps.Item(42).BasePointX, 0
        drw.Clamps.Item(42).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX + Xspace - drw.Clamps.Item(42).BasePointX, 0
    End If
    If drw.Clamps.Item(13).BasePointX - drw.Clamps.Item(10).BasePointX < Xspace Then
        drw.Clamps.Item(13).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(45).BasePointX, 0
        drw.Clamps.Item(14).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(45).BasePointX, 0
        drw.Clamps.Item(15).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(45).BasePointX, 0
        drw.Clamps.Item(28).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(45).BasePointX, 0
        drw.Clamps.Item(29).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(45).BasePointX, 0
        drw.Clamps.Item(30).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(45).BasePointX, 0
        drw.Clamps.Item(43).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(45).BasePointX, 0
        drw.Clamps.Item(44).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(45).BasePointX, 0
        drw.Clamps.Item(45).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(45).BasePointX, 0
    End If
    If Xend - drw.Clamps.Item(1).BasePointX < Xspace Then
        drw.Clamps.Item(1).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(33).BasePointX, 0
        drw.Clamps.Item(2).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(33).BasePointX, 0
        drw.Clamps.Item(3).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(33).BasePointX, 0
        drw.Clamps.Item(16).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(33).BasePointX, 0
        drw.Clamps.Item(17).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(33).BasePointX, 0
        drw.Clamps.Item(18).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(33).BasePointX, 0
        drw.Clamps.Item(31).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(33).BasePointX, 0
        drw.Clamps.Item(32).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(33).BasePointX, 0
        drw.Clamps.Item(33).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(33).BasePointX, 0
    End If


End If


End Sub
Public Sub bar3Accord2Move(i As Integer)
Dim drw As Drawing
Set drw = App.ActiveDrawing

Dim Xspace As Double
Xspace = 7.125
Dim Xend As Double


Xend = drw.Clamps.Item(i).BasePointX
If drw.Clamps.count = 30 Then

    If drw.Clamps.Item(10).BasePointX - drw.Clamps.Item(7).BasePointX < Xspace Then
        drw.Clamps.Item(10).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX + Xspace - drw.Clamps.Item(27).BasePointX, 0
        drw.Clamps.Item(11).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX + Xspace - drw.Clamps.Item(27).BasePointX, 0
        drw.Clamps.Item(12).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX + Xspace - drw.Clamps.Item(27).BasePointX, 0
        drw.Clamps.Item(25).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX + Xspace - drw.Clamps.Item(27).BasePointX, 0
        drw.Clamps.Item(26).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX + Xspace - drw.Clamps.Item(27).BasePointX, 0
        drw.Clamps.Item(27).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX + Xspace - drw.Clamps.Item(27).BasePointX, 0
    End If
    If drw.Clamps.Item(13).BasePointX - drw.Clamps.Item(10).BasePointX < Xspace Then
        drw.Clamps.Item(13).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(30).BasePointX, 0
        drw.Clamps.Item(14).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(30).BasePointX, 0
        drw.Clamps.Item(15).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(30).BasePointX, 0
        drw.Clamps.Item(28).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(30).BasePointX, 0
        drw.Clamps.Item(29).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(30).BasePointX, 0
        drw.Clamps.Item(30).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(30).BasePointX, 0
    End If
    If Xend - drw.Clamps.Item(4).BasePointX < Xspace Then
        drw.Clamps.Item(4).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(21).BasePointX, 0
        drw.Clamps.Item(5).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(21).BasePointX, 0
        drw.Clamps.Item(6).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(21).BasePointX, 0
        drw.Clamps.Item(19).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(21).BasePointX, 0
        drw.Clamps.Item(20).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(21).BasePointX, 0
        drw.Clamps.Item(21).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(21).BasePointX, 0
    End If
    If drw.Clamps.Item(4).BasePointX - drw.Clamps.Item(1).BasePointX < Xspace Then
        drw.Clamps.Item(1).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - Xspace - drw.Clamps.Item(18).BasePointX, 0
        drw.Clamps.Item(2).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - Xspace - drw.Clamps.Item(18).BasePointX, 0
        drw.Clamps.Item(3).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - Xspace - drw.Clamps.Item(18).BasePointX, 0
        drw.Clamps.Item(16).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - Xspace - drw.Clamps.Item(18).BasePointX, 0
        drw.Clamps.Item(17).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - Xspace - drw.Clamps.Item(18).BasePointX, 0
        drw.Clamps.Item(18).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - Xspace - drw.Clamps.Item(18).BasePointX, 0
    End If
    
ElseIf drw.Clamps.count = 45 Then

    If drw.Clamps.Item(10).BasePointX - drw.Clamps.Item(7).BasePointX < Xspace Then
        drw.Clamps.Item(10).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX + Xspace - drw.Clamps.Item(42).BasePointX, 0
        drw.Clamps.Item(11).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX + Xspace - drw.Clamps.Item(42).BasePointX, 0
        drw.Clamps.Item(12).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX + Xspace - drw.Clamps.Item(42).BasePointX, 0
        drw.Clamps.Item(25).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX + Xspace - drw.Clamps.Item(42).BasePointX, 0
        drw.Clamps.Item(26).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX + Xspace - drw.Clamps.Item(42).BasePointX, 0
        drw.Clamps.Item(27).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX + Xspace - drw.Clamps.Item(42).BasePointX, 0
        drw.Clamps.Item(40).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX + Xspace - drw.Clamps.Item(42).BasePointX, 0
        drw.Clamps.Item(41).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX + Xspace - drw.Clamps.Item(42).BasePointX, 0
        drw.Clamps.Item(42).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX + Xspace - drw.Clamps.Item(42).BasePointX, 0
    End If
    If drw.Clamps.Item(13).BasePointX - drw.Clamps.Item(10).BasePointX < Xspace Then
        drw.Clamps.Item(13).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(45).BasePointX, 0
        drw.Clamps.Item(14).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(45).BasePointX, 0
        drw.Clamps.Item(15).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(45).BasePointX, 0
        drw.Clamps.Item(28).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(45).BasePointX, 0
        drw.Clamps.Item(29).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(45).BasePointX, 0
        drw.Clamps.Item(30).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(45).BasePointX, 0
        drw.Clamps.Item(43).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(45).BasePointX, 0
        drw.Clamps.Item(44).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(45).BasePointX, 0
        drw.Clamps.Item(45).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(45).BasePointX, 0
    End If
    If Xend - drw.Clamps.Item(4).BasePointX < Xspace Then
        drw.Clamps.Item(4).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(5).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(6).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(19).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(20).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(21).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(34).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(35).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(36).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(36).BasePointX, 0
    End If
    If drw.Clamps.Item(4).BasePointX - drw.Clamps.Item(1).BasePointX < Xspace Then
        drw.Clamps.Item(1).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - Xspace - drw.Clamps.Item(33).BasePointX, 0
        drw.Clamps.Item(2).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - Xspace - drw.Clamps.Item(33).BasePointX, 0
        drw.Clamps.Item(3).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - Xspace - drw.Clamps.Item(33).BasePointX, 0
        drw.Clamps.Item(16).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - Xspace - drw.Clamps.Item(33).BasePointX, 0
        drw.Clamps.Item(17).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - Xspace - drw.Clamps.Item(33).BasePointX, 0
        drw.Clamps.Item(18).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - Xspace - drw.Clamps.Item(33).BasePointX, 0
        drw.Clamps.Item(31).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - Xspace - drw.Clamps.Item(33).BasePointX, 0
        drw.Clamps.Item(32).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - Xspace - drw.Clamps.Item(33).BasePointX, 0
        drw.Clamps.Item(33).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - Xspace - drw.Clamps.Item(33).BasePointX, 0
    End If
End If


End Sub
Public Sub bar4Accord2Move(i As Integer)
Dim drw As Drawing
Set drw = App.ActiveDrawing


Dim Xspace As Double
Xspace = 7.125
Dim Xend As Double


Xend = drw.Clamps.Item(i).BasePointX
If drw.Clamps.count = 30 Then

    If drw.Clamps.Item(13).BasePointX - drw.Clamps.Item(10).BasePointX < Xspace Then
        drw.Clamps.Item(13).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(30).BasePointX, 0
        drw.Clamps.Item(14).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(30).BasePointX, 0
        drw.Clamps.Item(15).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(30).BasePointX, 0
        drw.Clamps.Item(28).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(30).BasePointX, 0
        drw.Clamps.Item(29).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(30).BasePointX, 0
        drw.Clamps.Item(30).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(30).BasePointX, 0
    End If
    If Xend - drw.Clamps.Item(7).BasePointX < Xspace Then
        drw.Clamps.Item(7).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(24).BasePointX, 0
        drw.Clamps.Item(8).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(24).BasePointX, 0
        drw.Clamps.Item(9).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(24).BasePointX, 0
        drw.Clamps.Item(22).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(24).BasePointX, 0
        drw.Clamps.Item(23).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(24).BasePointX, 0
        drw.Clamps.Item(24).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(24).BasePointX, 0
    End If
    If drw.Clamps.Item(7).BasePointX - drw.Clamps.Item(4).BasePointX < Xspace Then
        drw.Clamps.Item(4).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX - Xspace - drw.Clamps.Item(21).BasePointX, 0
        drw.Clamps.Item(5).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX - Xspace - drw.Clamps.Item(21).BasePointX, 0
        drw.Clamps.Item(6).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX - Xspace - drw.Clamps.Item(21).BasePointX, 0
        drw.Clamps.Item(19).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX - Xspace - drw.Clamps.Item(21).BasePointX, 0
        drw.Clamps.Item(20).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX - Xspace - drw.Clamps.Item(21).BasePointX, 0
        drw.Clamps.Item(21).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX - Xspace - drw.Clamps.Item(21).BasePointX, 0
    End If
    If drw.Clamps.Item(4).BasePointX - drw.Clamps.Item(1).BasePointX < Xspace Then
        drw.Clamps.Item(1).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - Xspace - drw.Clamps.Item(18).BasePointX, 0
        drw.Clamps.Item(2).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - Xspace - drw.Clamps.Item(18).BasePointX, 0
        drw.Clamps.Item(3).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - Xspace - drw.Clamps.Item(18).BasePointX, 0
        drw.Clamps.Item(16).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - Xspace - drw.Clamps.Item(18).BasePointX, 0
        drw.Clamps.Item(17).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - Xspace - drw.Clamps.Item(18).BasePointX, 0
        drw.Clamps.Item(18).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - Xspace - drw.Clamps.Item(18).BasePointX, 0
    End If
    
ElseIf drw.Clamps.count = 45 Then

    If drw.Clamps.Item(13).BasePointX - drw.Clamps.Item(10).BasePointX < Xspace Then
        drw.Clamps.Item(13).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(45).BasePointX, 0
        drw.Clamps.Item(14).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(45).BasePointX, 0
        drw.Clamps.Item(15).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(45).BasePointX, 0
        drw.Clamps.Item(28).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(45).BasePointX, 0
        drw.Clamps.Item(29).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(45).BasePointX, 0
        drw.Clamps.Item(30).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(45).BasePointX, 0
        drw.Clamps.Item(43).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(45).BasePointX, 0
        drw.Clamps.Item(44).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(45).BasePointX, 0
        drw.Clamps.Item(45).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX + Xspace - drw.Clamps.Item(45).BasePointX, 0
    End If
    If Xend - drw.Clamps.Item(7).BasePointX < Xspace Then
        drw.Clamps.Item(7).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(39).BasePointX, 0
        drw.Clamps.Item(8).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(39).BasePointX, 0
        drw.Clamps.Item(9).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(39).BasePointX, 0
        drw.Clamps.Item(22).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(39).BasePointX, 0
        drw.Clamps.Item(23).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(39).BasePointX, 0
        drw.Clamps.Item(24).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(39).BasePointX, 0
        drw.Clamps.Item(37).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(39).BasePointX, 0
        drw.Clamps.Item(38).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(39).BasePointX, 0
        drw.Clamps.Item(39).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(39).BasePointX, 0
    End If
    If drw.Clamps.Item(7).BasePointX - drw.Clamps.Item(4).BasePointX < Xspace Then
        drw.Clamps.Item(4).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(5).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(6).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(19).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(20).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(21).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(34).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(35).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(36).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX - Xspace - drw.Clamps.Item(36).BasePointX, 0
    End If
    If drw.Clamps.Item(4).BasePointX - drw.Clamps.Item(1).BasePointX < Xspace Then
        drw.Clamps.Item(1).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - Xspace - drw.Clamps.Item(33).BasePointX, 0
        drw.Clamps.Item(2).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - Xspace - drw.Clamps.Item(33).BasePointX, 0
        drw.Clamps.Item(3).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - Xspace - drw.Clamps.Item(33).BasePointX, 0
        drw.Clamps.Item(16).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - Xspace - drw.Clamps.Item(33).BasePointX, 0
        drw.Clamps.Item(17).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - Xspace - drw.Clamps.Item(33).BasePointX, 0
        drw.Clamps.Item(18).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - Xspace - drw.Clamps.Item(33).BasePointX, 0
        drw.Clamps.Item(31).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - Xspace - drw.Clamps.Item(33).BasePointX, 0
        drw.Clamps.Item(32).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - Xspace - drw.Clamps.Item(33).BasePointX, 0
        drw.Clamps.Item(33).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - Xspace - drw.Clamps.Item(33).BasePointX, 0
    End If
    
End If

End Sub
Public Sub bar5Accord2Move(i As Integer)
Dim drw As Drawing
Set drw = App.ActiveDrawing

Dim Xspace As Double
Xspace = 7.125
Dim Xend As Double


Xend = drw.Clamps.Item(i).BasePointX
If drw.Clamps.count = 30 Then

    If Xend - drw.Clamps.Item(10).BasePointX < Xspace Then
        drw.Clamps.Item(10).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(27).BasePointX, 0
        drw.Clamps.Item(11).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(27).BasePointX, 0
        drw.Clamps.Item(12).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(27).BasePointX, 0
        drw.Clamps.Item(25).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(27).BasePointX, 0
        drw.Clamps.Item(26).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(27).BasePointX, 0
        drw.Clamps.Item(27).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(27).BasePointX, 0
    End If
    If drw.Clamps.Item(10).BasePointX - drw.Clamps.Item(7).BasePointX < Xspace Then
        drw.Clamps.Item(7).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX - Xspace - drw.Clamps.Item(24).BasePointX, 0
        drw.Clamps.Item(8).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX - Xspace - drw.Clamps.Item(24).BasePointX, 0
        drw.Clamps.Item(9).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX - Xspace - drw.Clamps.Item(24).BasePointX, 0
        drw.Clamps.Item(22).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX - Xspace - drw.Clamps.Item(24).BasePointX, 0
        drw.Clamps.Item(23).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX - Xspace - drw.Clamps.Item(24).BasePointX, 0
        drw.Clamps.Item(24).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX - Xspace - drw.Clamps.Item(24).BasePointX, 0
    End If
    If drw.Clamps.Item(7).BasePointX - drw.Clamps.Item(4).BasePointX < Xspace Then
        drw.Clamps.Item(4).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX - Xspace - drw.Clamps.Item(21).BasePointX, 0
        drw.Clamps.Item(5).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX - Xspace - drw.Clamps.Item(21).BasePointX, 0
        drw.Clamps.Item(6).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX - Xspace - drw.Clamps.Item(21).BasePointX, 0
        drw.Clamps.Item(19).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX - Xspace - drw.Clamps.Item(21).BasePointX, 0
        drw.Clamps.Item(20).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX - Xspace - drw.Clamps.Item(21).BasePointX, 0
        drw.Clamps.Item(21).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX - Xspace - drw.Clamps.Item(21).BasePointX, 0
    End If
    If drw.Clamps.Item(4).BasePointX - drw.Clamps.Item(1).BasePointX < Xspace Then
        drw.Clamps.Item(1).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - Xspace - drw.Clamps.Item(18).BasePointX, 0
        drw.Clamps.Item(2).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - Xspace - drw.Clamps.Item(18).BasePointX, 0
        drw.Clamps.Item(3).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - Xspace - drw.Clamps.Item(18).BasePointX, 0
        drw.Clamps.Item(16).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - Xspace - drw.Clamps.Item(18).BasePointX, 0
        drw.Clamps.Item(17).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - Xspace - drw.Clamps.Item(18).BasePointX, 0
        drw.Clamps.Item(18).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - Xspace - drw.Clamps.Item(18).BasePointX, 0
    End If

ElseIf drw.Clamps.count = 45 Then

    If Xend - drw.Clamps.Item(10).BasePointX < Xspace Then
        drw.Clamps.Item(10).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(42).BasePointX, 0
        drw.Clamps.Item(11).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(42).BasePointX, 0
        drw.Clamps.Item(12).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(42).BasePointX, 0
        drw.Clamps.Item(25).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(42).BasePointX, 0
        drw.Clamps.Item(26).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(42).BasePointX, 0
        drw.Clamps.Item(27).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(42).BasePointX, 0
        drw.Clamps.Item(40).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(42).BasePointX, 0
        drw.Clamps.Item(41).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(42).BasePointX, 0
        drw.Clamps.Item(42).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(42).BasePointX, 0
    End If
    If drw.Clamps.Item(10).BasePointX - drw.Clamps.Item(7).BasePointX < Xspace Then
        drw.Clamps.Item(7).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX - Xspace - drw.Clamps.Item(39).BasePointX, 0
        drw.Clamps.Item(8).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX - Xspace - drw.Clamps.Item(39).BasePointX, 0
        drw.Clamps.Item(9).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX - Xspace - drw.Clamps.Item(39).BasePointX, 0
        drw.Clamps.Item(22).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX - Xspace - drw.Clamps.Item(39).BasePointX, 0
        drw.Clamps.Item(23).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX - Xspace - drw.Clamps.Item(39).BasePointX, 0
        drw.Clamps.Item(24).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX - Xspace - drw.Clamps.Item(39).BasePointX, 0
        drw.Clamps.Item(37).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX - Xspace - drw.Clamps.Item(39).BasePointX, 0
        drw.Clamps.Item(38).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX - Xspace - drw.Clamps.Item(39).BasePointX, 0
        drw.Clamps.Item(39).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX - Xspace - drw.Clamps.Item(39).BasePointX, 0
    End If
    If drw.Clamps.Item(7).BasePointX - drw.Clamps.Item(4).BasePointX < Xspace Then
        drw.Clamps.Item(4).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(5).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(6).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(19).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(20).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(21).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(34).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(35).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(36).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX - Xspace - drw.Clamps.Item(36).BasePointX, 0
    End If
    If drw.Clamps.Item(4).BasePointX - drw.Clamps.Item(1).BasePointX < Xspace Then
        drw.Clamps.Item(1).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - Xspace - drw.Clamps.Item(33).BasePointX, 0
        drw.Clamps.Item(2).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - Xspace - drw.Clamps.Item(33).BasePointX, 0
        drw.Clamps.Item(3).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - Xspace - drw.Clamps.Item(33).BasePointX, 0
        drw.Clamps.Item(16).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - Xspace - drw.Clamps.Item(33).BasePointX, 0
        drw.Clamps.Item(17).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - Xspace - drw.Clamps.Item(33).BasePointX, 0
        drw.Clamps.Item(18).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - Xspace - drw.Clamps.Item(33).BasePointX, 0
        drw.Clamps.Item(31).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - Xspace - drw.Clamps.Item(33).BasePointX, 0
        drw.Clamps.Item(32).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - Xspace - drw.Clamps.Item(33).BasePointX, 0
        drw.Clamps.Item(33).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - Xspace - drw.Clamps.Item(33).BasePointX, 0
    End If

End If

End Sub



Public Sub checkBoundsAccord2()
Dim drw As Drawing
Set drw = App.ActiveDrawing

Dim Xmin As Double
Dim Ymin As Double
Dim Xmax As Double
Dim Ymax As Double

Xmin = 3.458
Xmax = 121.5
Ymin = 3.2
Ymax = 58.1

If drw.Clamps.Item(1).BasePointY < Ymin Or drw.Clamps.Item(4).BasePointY < Ymin Or drw.Clamps.Item(7).BasePointY < Ymin Or drw.Clamps.Item(10).BasePointY < Ymin Or drw.Clamps.Item(13).BasePointY < Ymin Then
    MsgBox "Row 1 out of bounds! Minimum is " & Ymin & " inches"
    drw.Undo
End If
If drw.Clamps.Item(3).BasePointY > Ymax Or drw.Clamps.Item(6).BasePointY > Ymax Or drw.Clamps.Item(9).BasePointY > Ymax Or drw.Clamps.Item(12).BasePointY > Ymax Or drw.Clamps.Item(15).BasePointY > Ymax Then
    MsgBox "Row 3 out of bounds! Maximum is " & Ymax & " inches"
    drw.Undo
End If
If drw.Clamps.Item(1).BasePointX < Xmin Then
    MsgBox "Bar 1 out of bounds! Minimum is " & Xmin & " inches"
    drw.Undo
End If
If drw.Clamps.Item(13).BasePointX > Xmax Then
    MsgBox "Bar 5 out of bounds! Minimum is " & Xmax & " inches"
    drw.Undo
End If

End Sub
Public Sub checkAllignAccord2()
Dim drw As Drawing
Set drw = App.ActiveDrawing

If drw.Clamps.count = 30 Then

    If drw.Clamps.Item(1).BasePointX <> drw.Clamps.Item(2).BasePointX Or drw.Clamps.Item(2).BasePointX <> drw.Clamps.Item(3).BasePointX Then
        drw.Clamps.Item(2).GeometryPath.MoveL drw.Clamps.Item(1).BasePointX - drw.Clamps.Item(2).BasePointX, 0
        drw.Clamps.Item(3).GeometryPath.MoveL drw.Clamps.Item(1).BasePointX - drw.Clamps.Item(3).BasePointX, 0
        drw.Clamps.Item(17).GeometryPath.MoveL drw.Clamps.Item(1).BasePointX - drw.Clamps.Item(17).BasePointX, 0
        drw.Clamps.Item(18).GeometryPath.MoveL drw.Clamps.Item(1).BasePointX - drw.Clamps.Item(18).BasePointX, 0
    End If
    If drw.Clamps.Item(4).BasePointX <> drw.Clamps.Item(5).BasePointX Or drw.Clamps.Item(5).BasePointX <> drw.Clamps.Item(6).BasePointX Then
        drw.Clamps.Item(5).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - drw.Clamps.Item(5).BasePointX, 0
        drw.Clamps.Item(6).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - drw.Clamps.Item(6).BasePointX, 0
        drw.Clamps.Item(20).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - drw.Clamps.Item(20).BasePointX, 0
        drw.Clamps.Item(21).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - drw.Clamps.Item(21).BasePointX, 0
    End If
    If drw.Clamps.Item(7).BasePointX <> drw.Clamps.Item(8).BasePointX Or drw.Clamps.Item(8).BasePointX <> drw.Clamps.Item(9).BasePointX Then
        drw.Clamps.Item(8).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX - drw.Clamps.Item(8).BasePointX, 0
        drw.Clamps.Item(9).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX - drw.Clamps.Item(9).BasePointX, 0
        drw.Clamps.Item(23).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX - drw.Clamps.Item(23).BasePointX, 0
        drw.Clamps.Item(24).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX - drw.Clamps.Item(24).BasePointX, 0
    End If
    If drw.Clamps.Item(10).BasePointX <> drw.Clamps.Item(11).BasePointX Or drw.Clamps.Item(11).BasePointX <> drw.Clamps.Item(12).BasePointX Then
        drw.Clamps.Item(11).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX - drw.Clamps.Item(11).BasePointX, 0
        drw.Clamps.Item(12).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX - drw.Clamps.Item(12).BasePointX, 0
        drw.Clamps.Item(26).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX - drw.Clamps.Item(26).BasePointX, 0
        drw.Clamps.Item(27).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX - drw.Clamps.Item(27).BasePointX, 0
    End If
    If drw.Clamps.Item(13).BasePointX <> drw.Clamps.Item(14).BasePointX Or drw.Clamps.Item(14).BasePointX <> drw.Clamps.Item(15).BasePointX Then
        drw.Clamps.Item(14).GeometryPath.MoveL drw.Clamps.Item(13).BasePointX - drw.Clamps.Item(14).BasePointX, 0
        drw.Clamps.Item(15).GeometryPath.MoveL drw.Clamps.Item(13).BasePointX - drw.Clamps.Item(15).BasePointX, 0
        drw.Clamps.Item(29).GeometryPath.MoveL drw.Clamps.Item(13).BasePointX - drw.Clamps.Item(29).BasePointX, 0
        drw.Clamps.Item(30).GeometryPath.MoveL drw.Clamps.Item(13).BasePointX - drw.Clamps.Item(30).BasePointX, 0
    End If

ElseIf drw.Clamps.count = 45 Then

    If drw.Clamps.Item(1).BasePointX <> drw.Clamps.Item(2).BasePointX Or drw.Clamps.Item(2).BasePointX <> drw.Clamps.Item(3).BasePointX Then
        drw.Clamps.Item(2).GeometryPath.MoveL drw.Clamps.Item(1).BasePointX - drw.Clamps.Item(2).BasePointX, 0
        drw.Clamps.Item(3).GeometryPath.MoveL drw.Clamps.Item(1).BasePointX - drw.Clamps.Item(3).BasePointX, 0
        drw.Clamps.Item(17).GeometryPath.MoveL drw.Clamps.Item(1).BasePointX - drw.Clamps.Item(17).BasePointX, 0
        drw.Clamps.Item(18).GeometryPath.MoveL drw.Clamps.Item(1).BasePointX - drw.Clamps.Item(18).BasePointX, 0
        drw.Clamps.Item(32).GeometryPath.MoveL drw.Clamps.Item(1).BasePointX - drw.Clamps.Item(32).BasePointX, 0
        drw.Clamps.Item(33).GeometryPath.MoveL drw.Clamps.Item(1).BasePointX - drw.Clamps.Item(33).BasePointX, 0
    End If
    If drw.Clamps.Item(4).BasePointX <> drw.Clamps.Item(5).BasePointX Or drw.Clamps.Item(5).BasePointX <> drw.Clamps.Item(6).BasePointX Then
        drw.Clamps.Item(5).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - drw.Clamps.Item(5).BasePointX, 0
        drw.Clamps.Item(6).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - drw.Clamps.Item(6).BasePointX, 0
        drw.Clamps.Item(20).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - drw.Clamps.Item(20).BasePointX, 0
        drw.Clamps.Item(21).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - drw.Clamps.Item(21).BasePointX, 0
        drw.Clamps.Item(35).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - drw.Clamps.Item(35).BasePointX, 0
        drw.Clamps.Item(36).GeometryPath.MoveL drw.Clamps.Item(4).BasePointX - drw.Clamps.Item(36).BasePointX, 0
    End If
    If drw.Clamps.Item(7).BasePointX <> drw.Clamps.Item(8).BasePointX Or drw.Clamps.Item(8).BasePointX <> drw.Clamps.Item(9).BasePointX Then
        drw.Clamps.Item(8).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX - drw.Clamps.Item(8).BasePointX, 0
        drw.Clamps.Item(9).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX - drw.Clamps.Item(9).BasePointX, 0
        drw.Clamps.Item(23).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX - drw.Clamps.Item(23).BasePointX, 0
        drw.Clamps.Item(24).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX - drw.Clamps.Item(24).BasePointX, 0
        drw.Clamps.Item(38).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX - drw.Clamps.Item(38).BasePointX, 0
        drw.Clamps.Item(39).GeometryPath.MoveL drw.Clamps.Item(7).BasePointX - drw.Clamps.Item(39).BasePointX, 0
    End If
    If drw.Clamps.Item(10).BasePointX <> drw.Clamps.Item(11).BasePointX Or drw.Clamps.Item(11).BasePointX <> drw.Clamps.Item(12).BasePointX Then
        drw.Clamps.Item(11).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX - drw.Clamps.Item(11).BasePointX, 0
        drw.Clamps.Item(12).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX - drw.Clamps.Item(12).BasePointX, 0
        drw.Clamps.Item(26).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX - drw.Clamps.Item(26).BasePointX, 0
        drw.Clamps.Item(27).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX - drw.Clamps.Item(27).BasePointX, 0
        drw.Clamps.Item(41).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX - drw.Clamps.Item(41).BasePointX, 0
        drw.Clamps.Item(42).GeometryPath.MoveL drw.Clamps.Item(10).BasePointX - drw.Clamps.Item(42).BasePointX, 0
    End If
    If drw.Clamps.Item(13).BasePointX <> drw.Clamps.Item(14).BasePointX Or drw.Clamps.Item(14).BasePointX <> drw.Clamps.Item(15).BasePointX Then
        drw.Clamps.Item(14).GeometryPath.MoveL drw.Clamps.Item(13).BasePointX - drw.Clamps.Item(14).BasePointX, 0
        drw.Clamps.Item(15).GeometryPath.MoveL drw.Clamps.Item(13).BasePointX - drw.Clamps.Item(15).BasePointX, 0
        drw.Clamps.Item(29).GeometryPath.MoveL drw.Clamps.Item(13).BasePointX - drw.Clamps.Item(29).BasePointX, 0
        drw.Clamps.Item(30).GeometryPath.MoveL drw.Clamps.Item(13).BasePointX - drw.Clamps.Item(30).BasePointX, 0
        drw.Clamps.Item(44).GeometryPath.MoveL drw.Clamps.Item(13).BasePointX - drw.Clamps.Item(44).BasePointX, 0
        drw.Clamps.Item(45).GeometryPath.MoveL drw.Clamps.Item(13).BasePointX - drw.Clamps.Item(45).BasePointX, 0
    End If
    
End If


End Sub
