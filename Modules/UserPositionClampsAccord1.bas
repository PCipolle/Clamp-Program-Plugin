Attribute VB_Name = "UserPositionClampsAccord1"
Public Sub moveRowsAccord1(i As Integer)

Dim drw As Drawing
Set drw = App.ActiveDrawing
Dim j As Integer
Dim k As Integer


    j = i + 16
    k = i + 32

Dim Yspace As Double
Yspace = 6.5
Dim Xend As Double
Dim Yend As Double

Xend = drw.Clamps.Item(i).BasePointX
Yend = drw.Clamps.Item(i).BasePointY

If drw.Clamps.count = 32 Then

    If i = 1 Or i = 5 Or i = 9 Or i = 13 Then
    
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
        If drw.Clamps.Item(i + 3).BasePointY - drw.Clamps.Item(i + 2).BasePointY < Yspace Then
            drw.Clamps.Item(i + 3).GeometryPath.MoveL Xend - drw.Clamps.Item(i + 3).BasePointX, drw.Clamps.Item(i + 2).BasePointY + Yspace - drw.Clamps.Item(i + 3).BasePointY
            drw.Clamps.Item(j + 3).GeometryPath.MoveL Xend - drw.Clamps.Item(j + 3).BasePointX, drw.Clamps.Item(j + 2).BasePointY + Yspace - drw.Clamps.Item(j + 3).BasePointY
        Else
            drw.Clamps.Item(i + 3).GeometryPath.MoveL Xend - drw.Clamps.Item(i + 3).BasePointX, 0
            drw.Clamps.Item(j + 3).GeometryPath.MoveL Xend - drw.Clamps.Item(j + 3).BasePointX, 0
        End If
    End If
    If i = 2 Or i = 6 Or i = 10 Or i = 14 Then
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
        If Yend - drw.Clamps.Item(i - 1).BasePointY < Yspace Then
            drw.Clamps.Item(i - 1).GeometryPath.MoveL Xend - drw.Clamps.Item(i - 1).BasePointX, Yend - Yspace - drw.Clamps.Item(i - 1).BasePointY
            drw.Clamps.Item(j - 1).GeometryPath.MoveL Xend - drw.Clamps.Item(j - 1).BasePointX, Yend - Yspace - drw.Clamps.Item(j - 1).BasePointY
        Else
            drw.Clamps.Item(i - 1).GeometryPath.MoveL Xend - drw.Clamps.Item(i - 1).BasePointX, 0
            drw.Clamps.Item(j - 1).GeometryPath.MoveL Xend - drw.Clamps.Item(j - 1).BasePointX, 0
        End If
    End If
    If i = 3 Or i = 7 Or i = 11 Or i = 15 Then
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
        If drw.Clamps.Item(i - 1).BasePointY - drw.Clamps.Item(i - 2).BasePointY < Yspace Then
            drw.Clamps.Item(i - 2).GeometryPath.MoveL Xend - drw.Clamps.Item(i - 2).BasePointX, drw.Clamps.Item(i - 1).BasePointY - Yspace - drw.Clamps.Item(i - 2).BasePointY
            drw.Clamps.Item(j - 2).GeometryPath.MoveL Xend - drw.Clamps.Item(j - 2).BasePointX, drw.Clamps.Item(i - 1).BasePointY - Yspace - drw.Clamps.Item(j - 2).BasePointY
        Else
            drw.Clamps.Item(i - 2).GeometryPath.MoveL Xend - drw.Clamps.Item(i - 2).BasePointX, 0
            drw.Clamps.Item(j - 2).GeometryPath.MoveL Xend - drw.Clamps.Item(j - 2).BasePointX, 0
        End If
    End If
    If i = 4 Or i = 8 Or i = 12 Or i = 16 Then
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
        If drw.Clamps.Item(i - 2).BasePointY - drw.Clamps.Item(i - 3).BasePointY < Yspace Then
            drw.Clamps.Item(i - 3).GeometryPath.MoveL Xend - drw.Clamps.Item(i - 3).BasePointX, drw.Clamps.Item(i - 2).BasePointY - Yspace - drw.Clamps.Item(i - 3).BasePointY
            drw.Clamps.Item(j - 3).GeometryPath.MoveL Xend - drw.Clamps.Item(j - 3).BasePointX, drw.Clamps.Item(i - 2).BasePointY - Yspace - drw.Clamps.Item(j - 3).BasePointY
        Else
            drw.Clamps.Item(i - 3).GeometryPath.MoveL Xend - drw.Clamps.Item(i - 3).BasePointX, 0
            drw.Clamps.Item(j - 3).GeometryPath.MoveL Xend - drw.Clamps.Item(j - 3).BasePointX, 0
        End If
    End If
    
ElseIf drw.Clamps.count = 48 Then

    If i = 1 Or i = 5 Or i = 9 Or i = 13 Then
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
        If drw.Clamps.Item(i + 3).BasePointY - drw.Clamps.Item(i + 2).BasePointY < Yspace Then
            drw.Clamps.Item(i + 3).GeometryPath.MoveL Xend - drw.Clamps.Item(i + 3).BasePointX, drw.Clamps.Item(i + 2).BasePointY + Yspace - drw.Clamps.Item(i + 3).BasePointY
            drw.Clamps.Item(j + 3).GeometryPath.MoveL Xend - drw.Clamps.Item(j + 3).BasePointX, drw.Clamps.Item(j + 2).BasePointY + Yspace - drw.Clamps.Item(j + 3).BasePointY
            drw.Clamps.Item(k + 3).GeometryPath.MoveL Xend - drw.Clamps.Item(k + 3).BasePointX, drw.Clamps.Item(k + 2).BasePointY + Yspace - drw.Clamps.Item(k + 3).BasePointY
        Else
            drw.Clamps.Item(i + 3).GeometryPath.MoveL Xend - drw.Clamps.Item(i + 3).BasePointX, 0
            drw.Clamps.Item(j + 3).GeometryPath.MoveL Xend - drw.Clamps.Item(j + 3).BasePointX, 0
            drw.Clamps.Item(k + 3).GeometryPath.MoveL Xend - drw.Clamps.Item(k + 3).BasePointX, 0
        End If
    End If
    If i = 2 Or i = 6 Or i = 10 Or i = 14 Then
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
    If i = 3 Or i = 7 Or i = 11 Or i = 15 Then
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
        If drw.Clamps.Item(i - 1).BasePointY - drw.Clamps.Item(i - 2).BasePointY < Yspace Then
            drw.Clamps.Item(i - 2).GeometryPath.MoveL Xend - drw.Clamps.Item(i - 2).BasePointX, drw.Clamps.Item(i - 1).BasePointY - Yspace - drw.Clamps.Item(i - 2).BasePointY
            drw.Clamps.Item(j - 2).GeometryPath.MoveL Xend - drw.Clamps.Item(j - 2).BasePointX, drw.Clamps.Item(i - 1).BasePointY - Yspace - drw.Clamps.Item(j - 2).BasePointY
            drw.Clamps.Item(k - 2).GeometryPath.MoveL Xend - drw.Clamps.Item(k - 2).BasePointX, drw.Clamps.Item(k - 1).BasePointY - Yspace - drw.Clamps.Item(k - 2).BasePointY
        Else
            drw.Clamps.Item(i - 2).GeometryPath.MoveL Xend - drw.Clamps.Item(i - 2).BasePointX, 0
            drw.Clamps.Item(j - 2).GeometryPath.MoveL Xend - drw.Clamps.Item(j - 2).BasePointX, 0
            drw.Clamps.Item(k - 2).GeometryPath.MoveL Xend - drw.Clamps.Item(k - 2).BasePointX, 0
        End If
    End If
    If i = 4 Or i = 8 Or i = 12 Or i = 16 Then
    
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
            drw.Clamps.Item(k - 2).GeometryPath.MoveL Xend - drw.Clamps.Item(k - 2).BasePointX, drw.Clamps.Item(k - 1).BasePointY - Yspace - drw.Clamps.Item(k - 2).BasePointY
        Else
            drw.Clamps.Item(i - 2).GeometryPath.MoveL Xend - drw.Clamps.Item(i - 2).BasePointX, 0
            drw.Clamps.Item(j - 2).GeometryPath.MoveL Xend - drw.Clamps.Item(j - 2).BasePointX, 0
            drw.Clamps.Item(k - 2).GeometryPath.MoveL Xend - drw.Clamps.Item(k - 2).BasePointX, 0
        End If
        If drw.Clamps.Item(i - 2).BasePointY - drw.Clamps.Item(i - 3).BasePointY < Yspace Then
            drw.Clamps.Item(i - 3).GeometryPath.MoveL Xend - drw.Clamps.Item(i - 3).BasePointX, drw.Clamps.Item(i - 2).BasePointY - Yspace - drw.Clamps.Item(i - 3).BasePointY
            drw.Clamps.Item(j - 3).GeometryPath.MoveL Xend - drw.Clamps.Item(j - 3).BasePointX, drw.Clamps.Item(i - 2).BasePointY - Yspace - drw.Clamps.Item(j - 3).BasePointY
            drw.Clamps.Item(k - 3).GeometryPath.MoveL Xend - drw.Clamps.Item(k - 3).BasePointX, drw.Clamps.Item(k - 2).BasePointY - Yspace - drw.Clamps.Item(k - 3).BasePointY
        Else
            drw.Clamps.Item(i - 3).GeometryPath.MoveL Xend - drw.Clamps.Item(i - 3).BasePointX, 0
            drw.Clamps.Item(j - 3).GeometryPath.MoveL Xend - drw.Clamps.Item(j - 3).BasePointX, 0
            drw.Clamps.Item(k - 3).GeometryPath.MoveL Xend - drw.Clamps.Item(k - 3).BasePointX, 0
        End If
    End If
End If

moveBarsAccord1 i

End Sub
Public Sub moveBarsAccord1(i As Integer)

If i = 1 Or i = 2 Or i = 3 Or i = 4 Then
    bar1Accord1Move i
End If
If i = 5 Or i = 6 Or i = 7 Or i = 8 Then
    bar2Accord1Move i
End If
If i = 9 Or i = 10 Or i = 11 Or i = 12 Then
    bar3Accord1Move i
End If
If i = 13 Or i = 14 Or i = 15 Or i = 16 Then
    bar4Accord1Move i
End If

End Sub
Public Sub bar1Accord1Move(i As Integer)
Dim drw As Drawing
Set drw = App.ActiveDrawing

Dim j As Integer
Dim k As Integer

Dim Xspace As Double
Xspace = 7.125
Dim Xend As Double


Xend = drw.Clamps.Item(i).BasePointX

If drw.Clamps.count = 32 Then

    If drw.Clamps.Item(5).BasePointX - Xend < Xspace Then
        drw.Clamps.Item(5).GeometryPath.MoveL Xend + Xspace - drw.Clamps.Item(24).BasePointX, 0
        drw.Clamps.Item(6).GeometryPath.MoveL Xend + Xspace - drw.Clamps.Item(24).BasePointX, 0
        drw.Clamps.Item(7).GeometryPath.MoveL Xend + Xspace - drw.Clamps.Item(24).BasePointX, 0
        drw.Clamps.Item(8).GeometryPath.MoveL Xend + Xspace - drw.Clamps.Item(24).BasePointX, 0
        drw.Clamps.Item(21).GeometryPath.MoveL Xend + Xspace - drw.Clamps.Item(24).BasePointX, 0
        drw.Clamps.Item(22).GeometryPath.MoveL Xend + Xspace - drw.Clamps.Item(24).BasePointX, 0
        drw.Clamps.Item(23).GeometryPath.MoveL Xend + Xspace - drw.Clamps.Item(24).BasePointX, 0
        drw.Clamps.Item(24).GeometryPath.MoveL Xend + Xspace - drw.Clamps.Item(24).BasePointX, 0
    End If
    If drw.Clamps.Item(9).BasePointX - drw.Clamps.Item(5).BasePointX < Xspace Then
        drw.Clamps.Item(9).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX + Xspace - drw.Clamps.Item(28).BasePointX, 0
        drw.Clamps.Item(10).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX + Xspace - drw.Clamps.Item(28).BasePointX, 0
        drw.Clamps.Item(11).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX + Xspace - drw.Clamps.Item(28).BasePointX, 0
        drw.Clamps.Item(12).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX + Xspace - drw.Clamps.Item(28).BasePointX, 0
        drw.Clamps.Item(25).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX + Xspace - drw.Clamps.Item(28).BasePointX, 0
        drw.Clamps.Item(26).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX + Xspace - drw.Clamps.Item(28).BasePointX, 0
        drw.Clamps.Item(27).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX + Xspace - drw.Clamps.Item(28).BasePointX, 0
        drw.Clamps.Item(28).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX + Xspace - drw.Clamps.Item(28).BasePointX, 0
    End If
    If drw.Clamps.Item(13).BasePointX - drw.Clamps.Item(9).BasePointX < Xspace Then
        drw.Clamps.Item(13).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(32).BasePointX, 0
        drw.Clamps.Item(14).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(32).BasePointX, 0
        drw.Clamps.Item(15).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(32).BasePointX, 0
        drw.Clamps.Item(16).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(32).BasePointX, 0
        drw.Clamps.Item(29).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(32).BasePointX, 0
        drw.Clamps.Item(30).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(32).BasePointX, 0
        drw.Clamps.Item(31).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(32).BasePointX, 0
        drw.Clamps.Item(32).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(32).BasePointX, 0
    End If

ElseIf drw.Clamps.count = 48 Then

    If drw.Clamps.Item(5).BasePointX - Xend < Xspace Then
        drw.Clamps.Item(5).GeometryPath.MoveL Xend + Xspace - drw.Clamps.Item(40).BasePointX, 0
        drw.Clamps.Item(6).GeometryPath.MoveL Xend + Xspace - drw.Clamps.Item(40).BasePointX, 0
        drw.Clamps.Item(7).GeometryPath.MoveL Xend + Xspace - drw.Clamps.Item(40).BasePointX, 0
        drw.Clamps.Item(8).GeometryPath.MoveL Xend + Xspace - drw.Clamps.Item(40).BasePointX, 0
        drw.Clamps.Item(21).GeometryPath.MoveL Xend + Xspace - drw.Clamps.Item(40).BasePointX, 0
        drw.Clamps.Item(22).GeometryPath.MoveL Xend + Xspace - drw.Clamps.Item(40).BasePointX, 0
        drw.Clamps.Item(23).GeometryPath.MoveL Xend + Xspace - drw.Clamps.Item(40).BasePointX, 0
        drw.Clamps.Item(24).GeometryPath.MoveL Xend + Xspace - drw.Clamps.Item(40).BasePointX, 0
        drw.Clamps.Item(37).GeometryPath.MoveL Xend + Xspace - drw.Clamps.Item(40).BasePointX, 0
        drw.Clamps.Item(38).GeometryPath.MoveL Xend + Xspace - drw.Clamps.Item(40).BasePointX, 0
        drw.Clamps.Item(39).GeometryPath.MoveL Xend + Xspace - drw.Clamps.Item(40).BasePointX, 0
        drw.Clamps.Item(40).GeometryPath.MoveL Xend + Xspace - drw.Clamps.Item(40).BasePointX, 0
    End If
    If drw.Clamps.Item(9).BasePointX - drw.Clamps.Item(5).BasePointX < Xspace Then
        drw.Clamps.Item(9).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX + Xspace - drw.Clamps.Item(44).BasePointX, 0
        drw.Clamps.Item(10).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX + Xspace - drw.Clamps.Item(44).BasePointX, 0
        drw.Clamps.Item(11).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX + Xspace - drw.Clamps.Item(44).BasePointX, 0
        drw.Clamps.Item(12).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX + Xspace - drw.Clamps.Item(44).BasePointX, 0
        drw.Clamps.Item(25).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX + Xspace - drw.Clamps.Item(44).BasePointX, 0
        drw.Clamps.Item(26).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX + Xspace - drw.Clamps.Item(44).BasePointX, 0
        drw.Clamps.Item(27).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX + Xspace - drw.Clamps.Item(44).BasePointX, 0
        drw.Clamps.Item(28).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX + Xspace - drw.Clamps.Item(44).BasePointX, 0
        drw.Clamps.Item(41).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX + Xspace - drw.Clamps.Item(44).BasePointX, 0
        drw.Clamps.Item(42).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX + Xspace - drw.Clamps.Item(44).BasePointX, 0
        drw.Clamps.Item(43).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX + Xspace - drw.Clamps.Item(44).BasePointX, 0
        drw.Clamps.Item(44).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX + Xspace - drw.Clamps.Item(44).BasePointX, 0
    End If
    If drw.Clamps.Item(13).BasePointX - drw.Clamps.Item(9).BasePointX < Xspace Then
        drw.Clamps.Item(13).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(48).BasePointX, 0
        drw.Clamps.Item(14).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(48).BasePointX, 0
        drw.Clamps.Item(15).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(48).BasePointX, 0
        drw.Clamps.Item(16).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(48).BasePointX, 0
        drw.Clamps.Item(29).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(48).BasePointX, 0
        drw.Clamps.Item(30).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(48).BasePointX, 0
        drw.Clamps.Item(31).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(48).BasePointX, 0
        drw.Clamps.Item(32).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(48).BasePointX, 0
        drw.Clamps.Item(45).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(48).BasePointX, 0
        drw.Clamps.Item(46).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(48).BasePointX, 0
        drw.Clamps.Item(47).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(48).BasePointX, 0
        drw.Clamps.Item(48).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(48).BasePointX, 0
    End If

End If


End Sub
Public Sub bar2Accord1Move(i As Integer)
Dim drw As Drawing
Set drw = App.ActiveDrawing

Dim j As Integer
Dim k As Integer

Dim Xspace As Double
Xspace = 7.125
Dim Xend As Double


Xend = drw.Clamps.Item(i).BasePointX
If drw.Clamps.count = 32 Then

    If drw.Clamps.Item(9).BasePointX - drw.Clamps.Item(5).BasePointX < Xspace Then
        drw.Clamps.Item(9).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX + Xspace - drw.Clamps.Item(28).BasePointX, 0
        drw.Clamps.Item(10).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX + Xspace - drw.Clamps.Item(28).BasePointX, 0
        drw.Clamps.Item(11).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX + Xspace - drw.Clamps.Item(28).BasePointX, 0
        drw.Clamps.Item(12).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX + Xspace - drw.Clamps.Item(28).BasePointX, 0
        drw.Clamps.Item(25).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX + Xspace - drw.Clamps.Item(28).BasePointX, 0
        drw.Clamps.Item(26).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX + Xspace - drw.Clamps.Item(28).BasePointX, 0
        drw.Clamps.Item(27).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX + Xspace - drw.Clamps.Item(28).BasePointX, 0
        drw.Clamps.Item(28).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX + Xspace - drw.Clamps.Item(28).BasePointX, 0
    End If
    If drw.Clamps.Item(13).BasePointX - drw.Clamps.Item(9).BasePointX < Xspace Then
        drw.Clamps.Item(13).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(32).BasePointX, 0
        drw.Clamps.Item(14).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(32).BasePointX, 0
        drw.Clamps.Item(15).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(32).BasePointX, 0
        drw.Clamps.Item(16).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(32).BasePointX, 0
        drw.Clamps.Item(29).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(32).BasePointX, 0
        drw.Clamps.Item(30).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(32).BasePointX, 0
        drw.Clamps.Item(31).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(32).BasePointX, 0
        drw.Clamps.Item(32).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(32).BasePointX, 0
    End If
    If Xend - drw.Clamps.Item(1).BasePointX < Xspace Then
        drw.Clamps.Item(1).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(20).BasePointX, 0
        drw.Clamps.Item(2).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(20).BasePointX, 0
        drw.Clamps.Item(3).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(20).BasePointX, 0
        drw.Clamps.Item(4).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(20).BasePointX, 0
        drw.Clamps.Item(17).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(20).BasePointX, 0
        drw.Clamps.Item(18).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(20).BasePointX, 0
        drw.Clamps.Item(19).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(20).BasePointX, 0
        drw.Clamps.Item(20).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(20).BasePointX, 0
    End If
    
ElseIf drw.Clamps.count = 48 Then

    If drw.Clamps.Item(9).BasePointX - drw.Clamps.Item(5).BasePointX < Xspace Then
        drw.Clamps.Item(9).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX + Xspace - drw.Clamps.Item(44).BasePointX, 0
        drw.Clamps.Item(10).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX + Xspace - drw.Clamps.Item(44).BasePointX, 0
        drw.Clamps.Item(11).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX + Xspace - drw.Clamps.Item(44).BasePointX, 0
        drw.Clamps.Item(12).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX + Xspace - drw.Clamps.Item(44).BasePointX, 0
        drw.Clamps.Item(25).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX + Xspace - drw.Clamps.Item(44).BasePointX, 0
        drw.Clamps.Item(26).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX + Xspace - drw.Clamps.Item(44).BasePointX, 0
        drw.Clamps.Item(27).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX + Xspace - drw.Clamps.Item(44).BasePointX, 0
        drw.Clamps.Item(28).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX + Xspace - drw.Clamps.Item(44).BasePointX, 0
        drw.Clamps.Item(41).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX + Xspace - drw.Clamps.Item(44).BasePointX, 0
        drw.Clamps.Item(42).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX + Xspace - drw.Clamps.Item(44).BasePointX, 0
        drw.Clamps.Item(43).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX + Xspace - drw.Clamps.Item(44).BasePointX, 0
        drw.Clamps.Item(44).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX + Xspace - drw.Clamps.Item(44).BasePointX, 0
    End If
    If drw.Clamps.Item(13).BasePointX - drw.Clamps.Item(9).BasePointX < Xspace Then
        drw.Clamps.Item(13).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(48).BasePointX, 0
        drw.Clamps.Item(14).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(48).BasePointX, 0
        drw.Clamps.Item(15).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(48).BasePointX, 0
        drw.Clamps.Item(16).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(48).BasePointX, 0
        drw.Clamps.Item(29).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(48).BasePointX, 0
        drw.Clamps.Item(30).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(48).BasePointX, 0
        drw.Clamps.Item(31).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(48).BasePointX, 0
        drw.Clamps.Item(32).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(48).BasePointX, 0
        drw.Clamps.Item(45).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(48).BasePointX, 0
        drw.Clamps.Item(46).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(48).BasePointX, 0
        drw.Clamps.Item(47).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(48).BasePointX, 0
        drw.Clamps.Item(48).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(48).BasePointX, 0
    End If
    If Xend - drw.Clamps.Item(1).BasePointX < Xspace Then
        drw.Clamps.Item(1).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(2).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(3).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(4).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(17).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(18).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(19).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(20).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(33).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(34).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(35).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(36).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(36).BasePointX, 0
    End If


End If


End Sub
Public Sub bar3Accord1Move(i As Integer)
Dim drw As Drawing
Set drw = App.ActiveDrawing

Dim j As Integer
Dim k As Integer

Dim Xspace As Double
Xspace = 7.125
Dim Xend As Double


Xend = drw.Clamps.Item(i).BasePointX

If drw.Clamps.count = 32 Then

    If drw.Clamps.Item(13).BasePointX - drw.Clamps.Item(9).BasePointX < Xspace Then
        drw.Clamps.Item(13).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(32).BasePointX, 0
        drw.Clamps.Item(14).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(32).BasePointX, 0
        drw.Clamps.Item(15).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(32).BasePointX, 0
        drw.Clamps.Item(16).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(32).BasePointX, 0
        drw.Clamps.Item(29).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(32).BasePointX, 0
        drw.Clamps.Item(30).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(32).BasePointX, 0
        drw.Clamps.Item(31).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(32).BasePointX, 0
        drw.Clamps.Item(32).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(32).BasePointX, 0
    End If
    If Xend - drw.Clamps.Item(5).BasePointX < Xspace Then
        drw.Clamps.Item(5).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(24).BasePointX, 0
        drw.Clamps.Item(6).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(24).BasePointX, 0
        drw.Clamps.Item(7).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(24).BasePointX, 0
        drw.Clamps.Item(8).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(24).BasePointX, 0
        drw.Clamps.Item(21).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(24).BasePointX, 0
        drw.Clamps.Item(22).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(24).BasePointX, 0
        drw.Clamps.Item(23).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(24).BasePointX, 0
        drw.Clamps.Item(24).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(24).BasePointX, 0
    End If
    If drw.Clamps.Item(5).BasePointX - drw.Clamps.Item(1).BasePointX < Xspace Then
        drw.Clamps.Item(1).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX - Xspace - drw.Clamps.Item(20).BasePointX, 0
        drw.Clamps.Item(2).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX - Xspace - drw.Clamps.Item(20).BasePointX, 0
        drw.Clamps.Item(3).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX - Xspace - drw.Clamps.Item(20).BasePointX, 0
        drw.Clamps.Item(4).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX - Xspace - drw.Clamps.Item(20).BasePointX, 0
        drw.Clamps.Item(17).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX - Xspace - drw.Clamps.Item(20).BasePointX, 0
        drw.Clamps.Item(18).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX - Xspace - drw.Clamps.Item(20).BasePointX, 0
        drw.Clamps.Item(19).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX - Xspace - drw.Clamps.Item(20).BasePointX, 0
        drw.Clamps.Item(20).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX - Xspace - drw.Clamps.Item(20).BasePointX, 0
    End If
    
ElseIf drw.Clamps.count = 48 Then

    If drw.Clamps.Item(13).BasePointX - drw.Clamps.Item(9).BasePointX < Xspace Then
        drw.Clamps.Item(13).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(48).BasePointX, 0
        drw.Clamps.Item(14).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(48).BasePointX, 0
        drw.Clamps.Item(15).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(48).BasePointX, 0
        drw.Clamps.Item(16).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(48).BasePointX, 0
        drw.Clamps.Item(29).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(48).BasePointX, 0
        drw.Clamps.Item(30).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(48).BasePointX, 0
        drw.Clamps.Item(31).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(48).BasePointX, 0
        drw.Clamps.Item(32).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(48).BasePointX, 0
        drw.Clamps.Item(45).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(48).BasePointX, 0
        drw.Clamps.Item(46).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(48).BasePointX, 0
        drw.Clamps.Item(47).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(48).BasePointX, 0
        drw.Clamps.Item(48).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX + Xspace - drw.Clamps.Item(48).BasePointX, 0
    End If
    If Xend - drw.Clamps.Item(5).BasePointX < Xspace Then
        drw.Clamps.Item(5).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(40).BasePointX, 0
        drw.Clamps.Item(6).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(40).BasePointX, 0
        drw.Clamps.Item(7).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(40).BasePointX, 0
        drw.Clamps.Item(8).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(40).BasePointX, 0
        drw.Clamps.Item(21).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(40).BasePointX, 0
        drw.Clamps.Item(22).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(40).BasePointX, 0
        drw.Clamps.Item(23).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(40).BasePointX, 0
        drw.Clamps.Item(24).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(40).BasePointX, 0
        drw.Clamps.Item(37).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(40).BasePointX, 0
        drw.Clamps.Item(38).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(40).BasePointX, 0
        drw.Clamps.Item(39).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(40).BasePointX, 0
        drw.Clamps.Item(40).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(40).BasePointX, 0
    End If
    If drw.Clamps.Item(5).BasePointX - drw.Clamps.Item(1).BasePointX < Xspace Then
        drw.Clamps.Item(1).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(2).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(3).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(4).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(17).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(18).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(19).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(20).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(33).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(34).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(35).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(36).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX - Xspace - drw.Clamps.Item(36).BasePointX, 0
    End If


End If


End Sub
Public Sub bar4Accord1Move(i As Integer)
Dim drw As Drawing
Set drw = App.ActiveDrawing

Dim j As Integer
Dim k As Integer

Dim Xspace As Double
Xspace = 7.125
Dim Xend As Double


Xend = drw.Clamps.Item(i).BasePointX

If drw.Clamps.count = 32 Then

    If Xend - drw.Clamps.Item(9).BasePointX < Xspace Then
        drw.Clamps.Item(9).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(28).BasePointX, 0
        drw.Clamps.Item(10).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(28).BasePointX, 0
        drw.Clamps.Item(11).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(28).BasePointX, 0
        drw.Clamps.Item(12).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(28).BasePointX, 0
        drw.Clamps.Item(25).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(28).BasePointX, 0
        drw.Clamps.Item(26).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(28).BasePointX, 0
        drw.Clamps.Item(27).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(28).BasePointX, 0
        drw.Clamps.Item(28).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(28).BasePointX, 0
    End If
    If drw.Clamps.Item(9).BasePointX - drw.Clamps.Item(5).BasePointX < Xspace Then
        drw.Clamps.Item(5).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX - Xspace - drw.Clamps.Item(24).BasePointX, 0
        drw.Clamps.Item(6).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX - Xspace - drw.Clamps.Item(24).BasePointX, 0
        drw.Clamps.Item(7).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX - Xspace - drw.Clamps.Item(24).BasePointX, 0
        drw.Clamps.Item(8).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX - Xspace - drw.Clamps.Item(24).BasePointX, 0
        drw.Clamps.Item(21).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX - Xspace - drw.Clamps.Item(24).BasePointX, 0
        drw.Clamps.Item(22).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX - Xspace - drw.Clamps.Item(24).BasePointX, 0
        drw.Clamps.Item(23).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX - Xspace - drw.Clamps.Item(24).BasePointX, 0
        drw.Clamps.Item(24).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX - Xspace - drw.Clamps.Item(24).BasePointX, 0
    End If
    If drw.Clamps.Item(5).BasePointX - drw.Clamps.Item(1).BasePointX < Xspace Then
        drw.Clamps.Item(1).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX - Xspace - drw.Clamps.Item(20).BasePointX, 0
        drw.Clamps.Item(2).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX - Xspace - drw.Clamps.Item(20).BasePointX, 0
        drw.Clamps.Item(3).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX - Xspace - drw.Clamps.Item(20).BasePointX, 0
        drw.Clamps.Item(4).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX - Xspace - drw.Clamps.Item(20).BasePointX, 0
        drw.Clamps.Item(17).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX - Xspace - drw.Clamps.Item(20).BasePointX, 0
        drw.Clamps.Item(18).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX - Xspace - drw.Clamps.Item(20).BasePointX, 0
        drw.Clamps.Item(19).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX - Xspace - drw.Clamps.Item(20).BasePointX, 0
        drw.Clamps.Item(20).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX - Xspace - drw.Clamps.Item(20).BasePointX, 0
    End If
    
ElseIf drw.Clamps.count = 48 Then

    If Xend - drw.Clamps.Item(9).BasePointX < Xspace Then
        drw.Clamps.Item(9).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(44).BasePointX, 0
        drw.Clamps.Item(10).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(44).BasePointX, 0
        drw.Clamps.Item(11).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(44).BasePointX, 0
        drw.Clamps.Item(12).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(44).BasePointX, 0
        drw.Clamps.Item(25).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(44).BasePointX, 0
        drw.Clamps.Item(26).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(44).BasePointX, 0
        drw.Clamps.Item(27).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(44).BasePointX, 0
        drw.Clamps.Item(28).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(44).BasePointX, 0
        drw.Clamps.Item(41).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(44).BasePointX, 0
        drw.Clamps.Item(42).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(44).BasePointX, 0
        drw.Clamps.Item(43).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(44).BasePointX, 0
        drw.Clamps.Item(44).GeometryPath.MoveL Xend - Xspace - drw.Clamps.Item(44).BasePointX, 0
    End If
    If drw.Clamps.Item(9).BasePointX - drw.Clamps.Item(5).BasePointX < Xspace Then
        drw.Clamps.Item(5).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX - Xspace - drw.Clamps.Item(40).BasePointX, 0
        drw.Clamps.Item(6).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX - Xspace - drw.Clamps.Item(40).BasePointX, 0
        drw.Clamps.Item(7).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX - Xspace - drw.Clamps.Item(40).BasePointX, 0
        drw.Clamps.Item(8).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX - Xspace - drw.Clamps.Item(40).BasePointX, 0
        drw.Clamps.Item(21).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX - Xspace - drw.Clamps.Item(40).BasePointX, 0
        drw.Clamps.Item(22).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX - Xspace - drw.Clamps.Item(40).BasePointX, 0
        drw.Clamps.Item(23).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX - Xspace - drw.Clamps.Item(40).BasePointX, 0
        drw.Clamps.Item(24).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX - Xspace - drw.Clamps.Item(40).BasePointX, 0
        drw.Clamps.Item(37).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX - Xspace - drw.Clamps.Item(40).BasePointX, 0
        drw.Clamps.Item(38).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX - Xspace - drw.Clamps.Item(40).BasePointX, 0
        drw.Clamps.Item(39).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX - Xspace - drw.Clamps.Item(40).BasePointX, 0
        drw.Clamps.Item(40).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX - Xspace - drw.Clamps.Item(40).BasePointX, 0
    End If
    If drw.Clamps.Item(5).BasePointX - drw.Clamps.Item(1).BasePointX < Xspace Then
        drw.Clamps.Item(1).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(2).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(3).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(4).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(17).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(18).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(19).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(20).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(33).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(34).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(35).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX - Xspace - drw.Clamps.Item(36).BasePointX, 0
        drw.Clamps.Item(36).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX - Xspace - drw.Clamps.Item(36).BasePointX, 0
    End If
    
End If

End Sub

Public Sub checkBoundsAccord1()
Dim drw As Drawing
Set drw = App.ActiveDrawing

Dim Xmin As Double
Dim Ymin As Double
Dim Xmax As Double
Dim Ymax As Double

Xmin = 3.25
Xmax = 96.5
Ymin = 3.4
Ymax = 58.26

If drw.Clamps.Item(1).BasePointY < Ymin Or drw.Clamps.Item(5).BasePointY < Ymin Or drw.Clamps.Item(9).BasePointY < Ymin Or drw.Clamps.Item(13).BasePointY < Ymin Then
    MsgBox "Row 1 out of bounds! Minimum is 3.4 inches"
    drw.Undo
End If
If drw.Clamps.Item(4).BasePointY > Ymax Or drw.Clamps.Item(8).BasePointY > Ymax Or drw.Clamps.Item(12).BasePointY > Ymax Or drw.Clamps.Item(16).BasePointY > Ymax Then
    MsgBox "Row 3 out of bounds! Maximum is 58.26 inches"
    drw.Undo
End If
If drw.Clamps.Item(1).BasePointX < Xmin Then
    MsgBox "Bar 1 out of bounds! Minimum is 3.25 inches"
    drw.Undo
End If
If drw.Clamps.Item(13).BasePointX > Xmax Then
    MsgBox "Bar 5 out of bounds! Minimum is 96.5 inches"
    drw.Undo
End If

End Sub
Public Sub checkAllignAccord1()
Dim drw As Drawing
Set drw = App.ActiveDrawing

If drw.Clamps.count = 32 Then

    If drw.Clamps.Item(1).BasePointX <> drw.Clamps.Item(2).BasePointX Or drw.Clamps.Item(2).BasePointX <> drw.Clamps.Item(3).BasePointX Or drw.Clamps.Item(3).BasePointX <> drw.Clamps.Item(4).BasePointX Then
        drw.Clamps.Item(2).GeometryPath.MoveL drw.Clamps.Item(1).BasePointX - drw.Clamps.Item(2).BasePointX, 0
        drw.Clamps.Item(3).GeometryPath.MoveL drw.Clamps.Item(1).BasePointX - drw.Clamps.Item(3).BasePointX, 0
        drw.Clamps.Item(4).GeometryPath.MoveL drw.Clamps.Item(1).BasePointX - drw.Clamps.Item(4).BasePointX, 0
        drw.Clamps.Item(18).GeometryPath.MoveL drw.Clamps.Item(1).BasePointX - drw.Clamps.Item(18).BasePointX, 0
        drw.Clamps.Item(19).GeometryPath.MoveL drw.Clamps.Item(1).BasePointX - drw.Clamps.Item(19).BasePointX, 0
        drw.Clamps.Item(20).GeometryPath.MoveL drw.Clamps.Item(1).BasePointX - drw.Clamps.Item(20).BasePointX, 0
    End If
    If drw.Clamps.Item(5).BasePointX <> drw.Clamps.Item(6).BasePointX Or drw.Clamps.Item(6).BasePointX <> drw.Clamps.Item(7).BasePointX Or drw.Clamps.Item(7).BasePointX <> drw.Clamps.Item(8).BasePointX Then
        drw.Clamps.Item(6).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX - drw.Clamps.Item(6).BasePointX, 0
        drw.Clamps.Item(7).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX - drw.Clamps.Item(7).BasePointX, 0
        drw.Clamps.Item(8).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX - drw.Clamps.Item(8).BasePointX, 0
        drw.Clamps.Item(22).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX - drw.Clamps.Item(22).BasePointX, 0
        drw.Clamps.Item(23).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX - drw.Clamps.Item(23).BasePointX, 0
        drw.Clamps.Item(24).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX - drw.Clamps.Item(24).BasePointX, 0
    End If
    If drw.Clamps.Item(9).BasePointX <> drw.Clamps.Item(10).BasePointX Or drw.Clamps.Item(10).BasePointX <> drw.Clamps.Item(11).BasePointX Or drw.Clamps.Item(11).BasePointX <> drw.Clamps.Item(12).BasePointX Then
        drw.Clamps.Item(10).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX - drw.Clamps.Item(10).BasePointX, 0
        drw.Clamps.Item(11).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX - drw.Clamps.Item(11).BasePointX, 0
        drw.Clamps.Item(12).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX - drw.Clamps.Item(12).BasePointX, 0
        drw.Clamps.Item(26).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX - drw.Clamps.Item(26).BasePointX, 0
        drw.Clamps.Item(27).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX - drw.Clamps.Item(27).BasePointX, 0
        drw.Clamps.Item(28).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX - drw.Clamps.Item(28).BasePointX, 0
    End If
    If drw.Clamps.Item(13).BasePointX <> drw.Clamps.Item(14).BasePointX Or drw.Clamps.Item(14).BasePointX <> drw.Clamps.Item(15).BasePointX Or drw.Clamps.Item(15).BasePointX <> drw.Clamps.Item(16).BasePointX Then
        drw.Clamps.Item(14).GeometryPath.MoveL drw.Clamps.Item(13).BasePointX - drw.Clamps.Item(14).BasePointX, 0
        drw.Clamps.Item(15).GeometryPath.MoveL drw.Clamps.Item(13).BasePointX - drw.Clamps.Item(15).BasePointX, 0
        drw.Clamps.Item(16).GeometryPath.MoveL drw.Clamps.Item(13).BasePointX - drw.Clamps.Item(16).BasePointX, 0
        drw.Clamps.Item(30).GeometryPath.MoveL drw.Clamps.Item(13).BasePointX - drw.Clamps.Item(30).BasePointX, 0
        drw.Clamps.Item(31).GeometryPath.MoveL drw.Clamps.Item(13).BasePointX - drw.Clamps.Item(31).BasePointX, 0
        drw.Clamps.Item(32).GeometryPath.MoveL drw.Clamps.Item(13).BasePointX - drw.Clamps.Item(32).BasePointX, 0
    End If
    
ElseIf drw.Clamps.count = 48 Then

    If drw.Clamps.Item(1).BasePointX <> drw.Clamps.Item(2).BasePointX Or drw.Clamps.Item(2).BasePointX <> drw.Clamps.Item(3).BasePointX Or drw.Clamps.Item(3).BasePointX <> drw.Clamps.Item(4).BasePointX Then
        drw.Clamps.Item(2).GeometryPath.MoveL drw.Clamps.Item(1).BasePointX - drw.Clamps.Item(2).BasePointX, 0
        drw.Clamps.Item(3).GeometryPath.MoveL drw.Clamps.Item(1).BasePointX - drw.Clamps.Item(3).BasePointX, 0
        drw.Clamps.Item(4).GeometryPath.MoveL drw.Clamps.Item(1).BasePointX - drw.Clamps.Item(4).BasePointX, 0
        drw.Clamps.Item(18).GeometryPath.MoveL drw.Clamps.Item(1).BasePointX - drw.Clamps.Item(18).BasePointX, 0
        drw.Clamps.Item(19).GeometryPath.MoveL drw.Clamps.Item(1).BasePointX - drw.Clamps.Item(19).BasePointX, 0
        drw.Clamps.Item(20).GeometryPath.MoveL drw.Clamps.Item(1).BasePointX - drw.Clamps.Item(20).BasePointX, 0
        drw.Clamps.Item(34).GeometryPath.MoveL drw.Clamps.Item(1).BasePointX - drw.Clamps.Item(34).BasePointX, 0
        drw.Clamps.Item(35).GeometryPath.MoveL drw.Clamps.Item(1).BasePointX - drw.Clamps.Item(35).BasePointX, 0
        drw.Clamps.Item(36).GeometryPath.MoveL drw.Clamps.Item(1).BasePointX - drw.Clamps.Item(36).BasePointX, 0
    End If
    If drw.Clamps.Item(5).BasePointX <> drw.Clamps.Item(6).BasePointX Or drw.Clamps.Item(6).BasePointX <> drw.Clamps.Item(7).BasePointX Or drw.Clamps.Item(7).BasePointX <> drw.Clamps.Item(8).BasePointX Then
        drw.Clamps.Item(6).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX - drw.Clamps.Item(6).BasePointX, 0
        drw.Clamps.Item(7).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX - drw.Clamps.Item(7).BasePointX, 0
        drw.Clamps.Item(8).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX - drw.Clamps.Item(8).BasePointX, 0
        drw.Clamps.Item(22).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX - drw.Clamps.Item(22).BasePointX, 0
        drw.Clamps.Item(23).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX - drw.Clamps.Item(23).BasePointX, 0
        drw.Clamps.Item(24).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX - drw.Clamps.Item(24).BasePointX, 0
        drw.Clamps.Item(38).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX - drw.Clamps.Item(38).BasePointX, 0
        drw.Clamps.Item(39).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX - drw.Clamps.Item(39).BasePointX, 0
        drw.Clamps.Item(40).GeometryPath.MoveL drw.Clamps.Item(5).BasePointX - drw.Clamps.Item(40).BasePointX, 0
    End If
    If drw.Clamps.Item(9).BasePointX <> drw.Clamps.Item(10).BasePointX Or drw.Clamps.Item(10).BasePointX <> drw.Clamps.Item(11).BasePointX Or drw.Clamps.Item(11).BasePointX <> drw.Clamps.Item(12).BasePointX Then
        drw.Clamps.Item(10).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX - drw.Clamps.Item(10).BasePointX, 0
        drw.Clamps.Item(11).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX - drw.Clamps.Item(11).BasePointX, 0
        drw.Clamps.Item(12).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX - drw.Clamps.Item(12).BasePointX, 0
        drw.Clamps.Item(26).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX - drw.Clamps.Item(26).BasePointX, 0
        drw.Clamps.Item(27).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX - drw.Clamps.Item(27).BasePointX, 0
        drw.Clamps.Item(28).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX - drw.Clamps.Item(28).BasePointX, 0
        drw.Clamps.Item(42).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX - drw.Clamps.Item(42).BasePointX, 0
        drw.Clamps.Item(43).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX - drw.Clamps.Item(43).BasePointX, 0
        drw.Clamps.Item(44).GeometryPath.MoveL drw.Clamps.Item(9).BasePointX - drw.Clamps.Item(44).BasePointX, 0
    End If
    If drw.Clamps.Item(13).BasePointX <> drw.Clamps.Item(14).BasePointX Or drw.Clamps.Item(14).BasePointX <> drw.Clamps.Item(15).BasePointX Or drw.Clamps.Item(15).BasePointX <> drw.Clamps.Item(16).BasePointX Then
        drw.Clamps.Item(14).GeometryPath.MoveL drw.Clamps.Item(13).BasePointX - drw.Clamps.Item(14).BasePointX, 0
        drw.Clamps.Item(15).GeometryPath.MoveL drw.Clamps.Item(13).BasePointX - drw.Clamps.Item(15).BasePointX, 0
        drw.Clamps.Item(16).GeometryPath.MoveL drw.Clamps.Item(13).BasePointX - drw.Clamps.Item(16).BasePointX, 0
        drw.Clamps.Item(30).GeometryPath.MoveL drw.Clamps.Item(13).BasePointX - drw.Clamps.Item(30).BasePointX, 0
        drw.Clamps.Item(31).GeometryPath.MoveL drw.Clamps.Item(13).BasePointX - drw.Clamps.Item(31).BasePointX, 0
        drw.Clamps.Item(32).GeometryPath.MoveL drw.Clamps.Item(13).BasePointX - drw.Clamps.Item(32).BasePointX, 0
        drw.Clamps.Item(46).GeometryPath.MoveL drw.Clamps.Item(13).BasePointX - drw.Clamps.Item(46).BasePointX, 0
        drw.Clamps.Item(47).GeometryPath.MoveL drw.Clamps.Item(13).BasePointX - drw.Clamps.Item(47).BasePointX, 0
        drw.Clamps.Item(48).GeometryPath.MoveL drw.Clamps.Item(13).BasePointX - drw.Clamps.Item(48).BasePointX, 0
    End If
    
End If


End Sub
