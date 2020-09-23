Attribute VB_Name = "Module2"
' Look to make 4: always prior
'Rw starts as 255; if rw<>255 then found something

Public Sub Easy() 'difficulty 0
' do 4 in a row
Look4Hor
If Rw <> 255 Then Exit Sub 'found to make 4 horizontal
Look4Vert
If Rw <> 255 Then Exit Sub 'found to make 4 vertical
Look4Dia1
If Rw <> 255 Then Exit Sub 'found to make 4 diagonal 1
Look4Dia2
End Sub

Public Sub NotSoEasy() 'difficulty 1
' do 4 in a row and defend
Look4Hor
If Rw <> 255 Then Exit Sub
Look4Vert
If Rw <> 255 Then Exit Sub
Look4Dia1
If Rw <> 255 Then Exit Sub
Look4Dia2
'look if player can make 4 in a row
LookPlayer4Hor
If Rw <> 255 Then Exit Sub
LookPlayer4Vert
If Rw <> 255 Then Exit Sub
LookPlayer4Dia1
If Rw <> 255 Then Exit Sub
LookPlayer4Dia2
If Rw <> 255 Then Exit Sub
End Sub

Public Sub Hard() 'difficulty 2
' do 4 in a row
Look4Hor
If Rw <> 255 Then Exit Sub
Look4Vert
If Rw <> 255 Then Exit Sub
Look4Dia1
If Rw <> 255 Then Exit Sub
Look4Dia2
'look if player can make 4 in a row
LookPlayer4Hor
If Rw <> 255 Then Exit Sub
LookPlayer4Vert
If Rw <> 255 Then Exit Sub
LookPlayer4Dia1
If Rw <> 255 Then Exit Sub
LookPlayer4Dia2
If Rw <> 255 Then Exit Sub
'do 3 in a row
Look3Hor
If Rw <> 255 Then Exit Sub
Look3Ver
If Rw <> 255 Then Exit Sub
Look3Dia1
End Sub

Public Sub VeryHard() 'difficulty 3
'do 4 in a row
Look4Hor
If Rw <> 255 Then Exit Sub 'found to make 4 horizontal
Look4Vert
If Rw <> 255 Then Exit Sub 'found to make 4 vertical
Look4Dia1
If Rw <> 255 Then Exit Sub 'found to make 4 diagonal 1
Look4Dia2
If Rw <> 255 Then Exit Sub 'found to make 4 diagonal 2
'----------------------------
'look if player can make 4 in a row
LookPlayer4Hor
If Rw <> 255 Then Exit Sub
LookPlayer4Vert
If Rw <> 255 Then Exit Sub
LookPlayer4Dia1
If Rw <> 255 Then Exit Sub
LookPlayer4Dia2
If Rw <> 255 Then Exit Sub
'----------------------------
'look if player can make 3 in a row
LookPlayer3Hor
If Rw <> 255 Then Exit Sub
LookPlayer3Vert
If Rw <> 255 Then Exit Sub
LookPlayer3Dia1
If Rw <> 255 Then Exit Sub
LookPlayer3Dia2
If Rw <> 255 Then Exit Sub
'----------------------------
'do 3 in a row
Look3Hor
If Rw <> 255 Then Exit Sub
Look3Ver
'----------------------------
End Sub
Public Sub ImpossibleToWin()
'do 4 in a row
Look4Hor
If Rw <> 255 Then Exit Sub 'found to make 4 horizontal
Look4Vert
If Rw <> 255 Then Exit Sub 'found to make 4 vertical
Look4Dia1
If Rw <> 255 Then Exit Sub 'found to make 4 diagonal 1
Look4Dia2
If Rw <> 255 Then Exit Sub 'found to make 4 diagonal 2
'----------------------------
'look if player can make 4 in a row
LookPlayer4Hor
If Rw <> 255 Then Exit Sub 'found player can make 4
LookPlayer4Vert
If Rw <> 255 Then Exit Sub
LookPlayer4Dia1
If Rw <> 255 Then Exit Sub
LookPlayer4Dia2
If Rw <> 255 Then Exit Sub
'----------------------------
'look if player can make 3 in a row
LookPlayer3Hor
If Rw <> 255 Then Exit Sub
LookPlayer3Vert
If Rw <> 255 Then Exit Sub
LookPlayer3Dia1
If Rw <> 255 Then Exit Sub
LookPlayer3Dia2
If Rw <> 255 Then Exit Sub
'----------------------------
'do 3 in a row
Look3Hor
If Rw <> 255 Then Exit Sub 'found to make 3 horizontal
Look3Ver
If Rw <> 255 Then Exit Sub 'found to make 3 vertical
Look3Dia1
If Rw <> 255 Then Exit Sub 'found to make 3 diagonal 1
Look3Dia2
If Rw <> 255 Then Exit Sub 'found to make 3 diagonal 2
'----------------------------
'do blocking
BlockPlayerHor
If Rw <> 255 Then Exit Sub
BlockPlayerVert
If Rw <> 255 Then Exit Sub
End Sub

Public Sub LookPlayer4Hor()
'DEFENCE !
'look if player can make 4 in a row horizontal
'only if place below <> 0
For xx = 0 To 6
For yy = 0 To 6
If yy < 6 Then
'look first
    If R(xx, yy) = 0 And R(xx, yy + 1) <> 0 And R(xx + 1, yy) = 1 And R(xx + 2, yy) = 1 And R(xx + 3, yy) = 1 Then
        Rw = xx
        Exit Sub
    End If
'look second
    If R(xx, yy) = 1 And R(xx + 1, yy) = 0 And R(xx + 1, yy + 1) <> 0 And R(xx + 2, yy) = 1 And R(xx + 3, yy) = 1 Then
        Rw = xx + 1
        Exit Sub
    End If
'look third
    If R(xx, yy) = 1 And R(xx + 1, yy) = 1 And R(xx + 2, yy) = 0 And R(xx + 2, yy + 1) <> 0 And R(xx + 3, yy) = 1 Then
        Rw = xx + 2
        Exit Sub
    End If
'look fourth
    If R(xx, yy) = 1 And R(xx + 1, yy) = 1 And R(xx + 2, yy) = 1 And R(xx + 3, yy) = 0 And R(xx + 3, yy + 1) <> 0 Then
        Rw = xx + 3
        Exit Sub
    End If
End If
'look if player can make 4 in a row horizontal
'on bottom row
If yy = 6 Then
'look first
    If R(xx, yy) = 0 And R(xx + 1, yy) = 1 And R(xx + 2, yy) = 1 And R(xx + 3, yy) = 1 Then
        Rw = xx
        Exit Sub
    End If
'look second
    If R(xx, yy) = 1 And R(xx + 1, yy) = 0 And R(xx + 2, yy) = 1 And R(xx + 3, yy) = 1 Then
        Rw = xx + 1
        Exit Sub
    End If
'look third
    If R(xx, yy) = 1 And R(xx + 1, yy) = 1 And R(xx + 2, yy) = 0 And R(xx + 3, yy) = 1 Then
        Rw = xx + 2
        Exit Sub
    End If
'look fourth
    If R(xx, yy) = 1 And R(xx + 1, yy) = 1 And R(xx + 2, yy) = 1 And R(xx + 3, yy) = 0 Then
        Rw = xx + 3
        Exit Sub
    End If
End If
Next yy
Next xx
End Sub

Public Sub LookPlayer4Vert() 'look if player can make 4 in a row vertically
For xx = 0 To 9
For yy = 0 To 3
If R(xx, yy) = 0 And R(xx, yy + 1) = 1 And R(xx, yy + 2) = 1 And R(xx, yy + 3) = 1 Then
    Rw = xx
    Exit Sub
End If
Next yy
Next xx
End Sub

Public Sub LookPlayer4Dia1()
'look if player can make 4 in a row diagonal 1
'only if place below <> 0
For xx = 0 To 6
For yy = 0 To 3
'look first
    If R(xx, yy) = 0 And R(xx, yy + 1) <> 0 And R(xx + 1, yy + 1) = 1 And R(xx + 2, yy + 2) = 1 And R(xx + 3, yy + 3) = 1 Then
        Rw = xx
        Exit Sub
    End If
'look second
    If R(xx, yy) = 1 And R(xx + 1, yy + 1) = 0 And R(xx + 1, yy + 2) <> 0 And R(xx + 2, yy + 2) = 1 And R(xx + 3, yy + 3) = 1 Then
        Rw = xx + 1
        Exit Sub
    End If
'look third
    If R(xx, yy) = 1 And R(xx + 1, yy + 1) = 1 And R(xx + 2, yy + 2) = 0 And R(xx + 2, yy + 3) <> 0 And R(xx + 3, yy + 3) = 1 Then
        Rw = xx + 2
        Exit Sub
    End If
'look fourth
If yy < 3 Then
    If R(xx, yy) = 1 And R(xx + 1, yy + 1) = 1 And R(xx + 2, yy + 2) = 1 And R(xx + 3, yy + 3) = 0 And R(xx + 3, yy + 4) <> 0 Then
        Rw = xx + 3
        Exit Sub
    End If
End If
'look fourth (2e chance)
If yy = 3 Then
    If R(xx, yy) = 1 And R(xx + 1, yy + 1) = 1 And R(xx + 2, yy + 2) = 1 And R(xx + 3, yy + 3) = 0 Then
        Rw = xx + 3
        Exit Sub
    End If
End If
Next yy
Next xx
End Sub

Public Sub LookPlayer4Dia2()
'look if player can make 4 in a row diagonal 2
'only if place below <> 0
For xx = 3 To 9
For yy = 0 To 3
'look first
If R(xx, yy) = 0 And R(xx, yy + 1) <> 0 And R(xx - 1, yy + 1) = 1 And R(xx - 2, yy + 2) = 1 And R(xx - 3, yy + 3) = 1 Then
    Rw = xx
    Exit Sub
End If
'look second
If R(xx, yy) = 1 And R(xx - 1, yy + 1) = 0 And R(xx - 1, yy + 2) <> 0 And R(xx - 2, yy + 2) = 1 And R(xx - 3, yy + 3) = 1 Then
    Rw = xx - 1
    Exit Sub
End If
'look third
If R(xx, yy) = 1 And R(xx - 1, yy + 1) = 1 And R(xx - 2, yy + 2) = 0 And R(xx - 2, yy + 3) <> 0 And R(xx - 3, yy + 3) = 1 Then
    Rw = xx - 2
    Exit Sub
End If
'look fourth
If yy < 3 Then
    If R(xx, yy) = 1 And R(xx - 1, yy + 1) = 1 And R(xx - 2, yy + 2) = 1 And R(xx - 3, yy + 3) = 0 And R(xx - 3, yy + 4) <> 0 Then
        Rw = xx - 3
        Exit Sub
    End If
End If
'look fourth (2e chance)
If yy = 3 Then
    If R(xx, yy) = 1 And R(xx - 1, yy + 1) = 1 And R(xx - 2, yy + 2) = 1 And R(xx - 3, yy + 3) = 0 Then
        Rw = xx - 3
        Exit Sub
    End If
End If
Next yy
Next xx
End Sub

Public Sub Look4Hor()
'look to make 4 in a row horizontal
'only if place below <> 0
For xx = 0 To 6
For yy = 0 To 6
If yy < 6 Then
'look first
    If R(xx, yy) = 0 And R(xx, yy + 1) <> 0 And R(xx + 1, yy) = 2 And R(xx + 2, yy) = 2 And R(xx + 3, yy) = 2 Then
        Rw = xx
        Exit Sub
    End If
'look second
    If R(xx, yy) = 2 And R(xx + 1, yy) = 0 And R(xx + 1, yy + 1) <> 0 And R(xx + 2, yy) = 2 And R(xx + 3, yy) = 2 Then
        Rw = xx + 1
        Exit Sub
    End If
'look third
    If R(xx, yy) = 2 And R(xx + 1, yy) = 2 And R(xx + 2, yy) = 0 And R(xx + 2, yy + 1) <> 0 And R(xx + 3, yy) = 2 Then
        Rw = xx + 2
        Exit Sub
    End If
'look fourth
    If R(xx, yy) = 2 And R(xx + 1, yy) = 2 And R(xx + 2, yy) = 2 And R(xx + 3, yy) = 0 And R(xx + 3, yy + 1) <> 0 Then
        Rw = xx + 3
        Exit Sub
    End If
End If
'look to make 4 in a row horizontal
'on bottom row
If yy = 6 Then
'look first
    If R(xx, yy) = 0 And R(xx + 1, yy) = 2 And R(xx + 2, yy) = 2 And R(xx + 3, yy) = 2 Then
        Rw = xx
        Exit Sub
    End If
'look second
    If R(xx, yy) = 2 And R(xx + 1, yy) = 0 And R(xx + 2, yy) = 2 And R(xx + 3, yy) = 2 Then
        Rw = xx + 1
        Exit Sub
    End If
'look third
    If R(xx, yy) = 2 And R(xx + 1, yy) = 2 And R(xx + 2, yy) = 0 And R(xx + 3, yy) = 2 Then
        Rw = xx + 2
        Exit Sub
    End If
'look fourth
    If R(xx, yy) = 2 And R(xx + 1, yy) = 2 And R(xx + 2, yy) = 2 And R(xx + 3, yy) = 0 Then
        Rw = xx + 3
        Exit Sub
    End If
End If
Next yy
Next xx
End Sub

Public Sub Look4Vert() 'look to make 4 in a row vertically
For xx = 0 To 9
For yy = 0 To 3
If R(xx, yy) = 0 And R(xx, yy + 1) = 2 And R(xx, yy + 2) = 2 And R(xx, yy + 3) = 2 Then
    Rw = xx
    Exit Sub
End If
Next yy
Next xx
End Sub

Public Sub Look4Dia1()
'look to make 4 in a row diagonal 1
'only if place below <> 0
For xx = 0 To 6
For yy = 0 To 3
'look first
    If R(xx, yy) = 0 And R(xx, yy + 1) <> 0 And R(xx + 1, yy + 1) = 2 And R(xx + 2, yy + 2) = 2 And R(xx + 3, yy + 3) = 2 Then
        Rw = xx
        Exit Sub
    End If
'look second
    If R(xx, yy) = 2 And R(xx + 1, yy + 1) = 0 And R(xx + 1, yy + 2) <> 0 And R(xx + 2, yy + 2) = 2 And R(xx + 3, yy + 3) = 2 Then
        Rw = xx + 1
        Exit Sub
    End If
'look third
    If R(xx, yy) = 2 And R(xx + 1, yy + 1) = 2 And R(xx + 2, yy + 2) = 0 And R(xx + 2, yy + 3) <> 0 And R(xx + 3, yy + 3) = 2 Then
        Rw = xx + 2
        Exit Sub
    End If
'look fourth
If yy < 3 Then
    If R(xx, yy) = 2 And R(xx + 1, yy + 1) = 2 And R(xx + 2, yy + 2) = 2 And R(xx + 3, yy + 3) = 0 And R(xx + 3, yy + 4) <> 0 Then
        Rw = xx + 3
        Exit Sub
    End If
End If
'look fourth (2e chance)
If yy = 3 Then
    If R(xx, yy) = 2 And R(xx + 1, yy + 1) = 2 And R(xx + 2, yy + 2) = 2 And R(xx + 3, yy + 3) = 0 Then
        Rw = xx + 3
        Exit Sub
    End If
End If
Next yy
Next xx
End Sub

Public Sub Look4Dia2()
'look to make 4 in a row diagonal 2
'only if place below <> 0
For xx = 3 To 9
For yy = 0 To 3
'look first
If R(xx, yy) = 0 And R(xx, yy + 1) <> 0 And R(xx - 1, yy + 1) = 2 And R(xx - 2, yy + 2) = 2 And R(xx - 3, yy + 3) = 2 Then
    Rw = xx
    Exit Sub
End If
'look second
If R(xx, yy) = 2 And R(xx - 1, yy + 1) = 0 And R(xx - 1, yy + 2) <> 0 And R(xx - 2, yy + 2) = 2 And R(xx - 3, yy + 3) = 2 Then
    Rw = xx - 1
    Exit Sub
End If
'look third
If R(xx, yy) = 2 And R(xx - 1, yy + 1) = 2 And R(xx - 2, yy + 2) = 0 And R(xx - 2, yy + 3) <> 0 And R(xx - 3, yy + 3) = 2 Then
    Rw = xx - 2
    Exit Sub
End If
'look fourth
If yy < 3 Then
    If R(xx, yy) = 2 And R(xx - 1, yy + 1) = 2 And R(xx - 2, yy + 2) = 2 And R(xx - 3, yy + 3) = 0 And R(xx - 3, yy + 4) <> 0 Then
        Rw = xx - 3
        Exit Sub
    End If
End If
'look fourth (2e chance)
If yy = 3 Then
    If R(xx, yy) = 2 And R(xx - 1, yy + 1) = 2 And R(xx - 2, yy + 2) = 2 And R(xx - 3, yy + 3) = 0 Then
        Rw = xx - 3
        Exit Sub
    End If
End If
Next yy
Next xx
End Sub

Public Sub Look3Hor()
'look to make 3 in a row horizontal
' if place below <> 0
For xx = 0 To 6
For yy = 0 To 6
If yy < 6 Then
'look first & second
If R(xx, yy) = 0 And R(xx, yy + 1) <> 0 And R(xx + 1, yy) = 0 And R(xx + 2, yy) = 2 And R(xx + 3, yy) = 2 Then
    Rw = xx
    Exit Sub
End If
'look first & second (2)
If R(xx, yy) = 0 And R(xx + 1, yy) = 0 And R(xx + 1, yy + 1) <> 0 And R(xx + 2, yy) = 2 And R(xx + 3, yy) = 2 Then
    Rw = xx + 1
    Exit Sub
End If
'look first & third
If R(xx, yy) = 0 And R(xx, yy + 1) <> 0 And R(xx + 1, yy) = 2 And R(xx + 2, yy) = 0 And R(xx + 3, yy) = 2 Then
    Rw = xx
    Exit Sub
End If
'look first & third (2)
If R(xx, yy) = 0 And R(xx + 1, yy) = 2 And R(xx + 2, yy) = 0 And R(xx + 2, yy + 1) <> 0 And R(xx + 3, yy) = 2 Then
    Rw = xx + 2
    Exit Sub
End If
'look first & fourth
If R(xx, yy) = 0 And R(xx, yy + 1) <> 0 And R(xx + 1, yy) = 2 And R(xx + 2, yy) = 2 And R(xx + 3, yy) = 0 Then
    Rw = xx
    Exit Sub
End If
'look first & fourth (2)
If R(xx, yy) = 0 And R(xx + 1, yy) = 2 And R(xx + 2, yy) = 2 And R(xx + 3, yy) = 0 And R(xx + 3, yy + 1) <> 0 Then
    Rw = xx + 3
    Exit Sub
End If
'look second & third
If R(xx, yy) = 2 And R(xx + 1, yy) = 0 And R(xx + 1, yy + 1) <> 0 And R(xx + 2, yy) = 0 And R(xx + 3, yy) = 2 Then
    Rw = xx + 1
    Exit Sub
End If
'look second & third (2)
If R(xx, yy) = 2 And R(xx + 1, yy) = 0 And R(xx + 2, yy) = 0 And R(xx + 2, yy + 1) <> 0 And R(xx + 3, yy) = 2 Then
    Rw = xx + 2
    Exit Sub
End If
'look second & fourth
If R(xx, yy) = 2 And R(xx + 1, yy) = 0 And R(xx + 1, yy + 1) <> 0 And R(xx + 2, yy) = 2 And R(xx + 3, yy) = 0 Then
    Rw = xx + 1
    Exit Sub
End If
'look second & fourth (2)
If R(xx, yy) = 2 And R(xx + 1, yy) = 0 And R(xx + 2, yy) = 2 And R(xx + 3, yy) = 0 And R(xx + 3, yy + 1) <> 0 Then
    Rw = xx + 3
    Exit Sub
End If
'look third & fourth
If R(xx, yy) = 2 And R(xx + 1, yy) = 2 And R(xx + 2, yy) = 0 And R(xx + 2, yy + 1) <> 0 And R(xx + 3, yy) = 0 Then
    Rw = xx + 2
    Exit Sub
End If
'look third & fourth (2)
If R(xx, yy) = 2 And R(xx + 1, yy) = 2 And R(xx + 2, yy) = 0 And R(xx + 3, yy) = 0 And R(xx + 3, yy + 1) <> 0 Then
    Rw = xx + 3
    Exit Sub
End If
End If

If yy = 6 Then
'look first & second
If R(xx, yy) = 0 And R(xx + 1, yy) = 0 And R(xx + 2, yy) = 2 And R(xx + 3, yy) = 2 Then
    Rw = xx
    Exit Sub
End If
'look first & third
If R(xx, yy) = 0 And R(xx + 1, yy) = 2 And R(xx + 2, yy) = 0 And R(xx + 3, yy) = 2 Then
    Rw = xx
    Exit Sub
End If
'look first & fourth
If R(xx, yy) = 0 And R(xx + 1, yy) = 2 And R(xx + 2, yy) = 2 And R(xx + 3, yy) = 0 Then
    Rw = xx
    Exit Sub
End If
'look second & third
If R(xx, yy) = 2 And R(xx + 1, yy) = 0 And R(xx + 2, yy) = 0 And R(xx + 3, yy) = 2 Then
    Rw = xx + 1
    Exit Sub
End If
'look second & fourth
If R(xx, yy) = 2 And R(xx + 1, yy) = 0 And R(xx + 2, yy) = 2 And R(xx + 3, yy) = 0 Then
    Rw = xx + 1
    Exit Sub
End If
'look third & fourth
If R(xx, yy) = 2 And R(xx + 1, yy) = 2 And R(xx + 2, yy) = 0 And R(xx + 3, yy) = 0 Then
    Rw = xx + 2
    Exit Sub
End If
End If
Next yy
Next xx
End Sub

Public Sub Look3Ver() 'look to make 3 in a row vertically
For xx = 0 To 9
For yy = 0 To 3
If R(xx, yy) = 0 And R(xx, yy + 1) = 0 And R(xx, yy + 2) = 2 And R(xx, yy + 3) = 2 Then
    Rw = xx
    Exit Sub
End If
Next yy
Next xx
End Sub

Public Sub Look3Dia1() 'look to make 3 in a row diagonal 1
For xx = 0 To 6
For yy = 0 To 3
'look first & second (first)
If R(xx, yy) = 0 And R(xx, yy + 1) <> 0 And R(xx + 1, yy + 1) = 0 And R(xx + 2, yy + 2) = 2 And R(xx + 3, yy + 3) = 2 Then
    Rw = xx
    Exit Sub
End If
'look first & second (second)
If R(xx, yy) = 0 And R(xx + 1, yy + 1) = 0 And R(xx + 1, yy + 2) <> 0 And R(xx + 2, yy + 2) = 2 And R(xx + 3, yy + 3) = 2 Then
    Rw = xx + 1
    Exit Sub
End If
'look first & third (first)
If R(xx, yy) = 0 And R(xx, yy + 1) <> 0 And R(xx + 1, yy + 1) = 2 And R(xx + 2, yy + 2) = 0 And R(xx + 3, yy + 3) = 2 Then
    Rw = xx
    Exit Sub
End If
'look first & third (third)
If R(xx, yy) = 0 And R(xx + 1, yy + 1) = 2 And R(xx + 2, yy + 2) = 0 And R(xx + 2, yy + 3) <> 0 And R(xx + 2, yy + 3) <> 0 And R(xx + 3, yy + 3) = 2 Then
    Rw = xx + 2
    Exit Sub
End If
'look first & fourth (first)
If R(xx, yy) = 0 And R(xx, yy + 1) <> 0 And R(xx + 1, yy + 1) = 2 And R(xx + 2, yy + 2) = 2 And R(xx + 3, yy + 3) = 0 Then
    Rw = xx
    Exit Sub
End If
'look first & fourth (fourth)
If yy < 3 Then
    If R(xx, yy) = 0 And R(xx + 1, yy + 1) = 2 And R(xx + 2, yy + 2) = 2 And R(xx + 3, yy + 3) = 0 And R(xx + 3, yy + 4) <> 0 Then
        Rw = xx + 3
        Exit Sub
    End If
End If
'look first & fourth (fourth)-2
If yy = 3 Then
    If R(xx, yy) = 0 And R(xx + 1, yy + 1) = 2 And R(xx + 2, yy + 2) = 2 And R(xx + 3, yy + 3) = 0 Then
        Rw = xx + 3
        Exit Sub
    End If
End If
'look second & third (second)
    If R(xx, yy) = 2 And R(xx + 1, yy + 1) = 0 And R(xx + 1, yy + 2) <> 0 And R(xx + 2, yy + 2) = 0 And R(xx + 3, yy + 3) = 2 Then
        Rw = xx + 1
        Exit Sub
    End If
'look second & third (third)
    If R(xx, yy) = 2 And R(xx + 1, yy + 1) = 0 And R(xx + 2, yy + 2) = 0 And R(xx + 2, yy + 3) <> 0 And R(xx + 3, yy + 3) = 2 Then
        Rw = xx + 2
        Exit Sub
    End If
'look second & fourth (second)
    If R(xx, yy) = 2 And R(xx + 1, yy + 1) = 0 And R(xx + 1, yy + 2) <> 0 And R(xx + 2, yy + 2) = 2 And R(xx + 3, yy + 3) = 0 Then
        Rw = xx + 1
        Exit Sub
    End If
'look second & fourth (fourth)
If yy < 3 Then
    If R(xx, yy) = 2 And R(xx + 1, yy + 1) = 0 And R(xx + 2, yy + 2) = 2 And R(xx + 3, yy + 3) = 0 And R(xx + 3, yy + 4) <> 0 Then
        Rw = xx + 3
        Exit Sub
    End If
End If
 'look second & fourth (fourth)-2
If yy = 3 Then
    If R(xx, yy) = 2 And R(xx + 1, yy + 1) = 0 And R(xx + 2, yy + 2) = 2 And R(xx + 3, yy + 3) = 0 Then
        Rw = xx + 3
        Exit Sub
    End If
End If
'look third & fourth (third)
    If R(xx, yy) = 2 And R(xx + 1, yy + 1) = 2 And R(xx + 2, yy + 2) = 0 And R(xx + 2, yy + 3) <> 0 And R(xx + 3, yy + 3) = 0 Then
        Rw = xx + 2
        Exit Sub
    End If
'look third & fourth (fourth)
If yy < 3 Then
    If R(xx, yy) = 2 And R(xx + 1, yy + 1) = 2 And R(xx + 2, yy + 2) = 0 And R(xx + 3, yy + 3) = 0 And R(xx + 3, yy + 4) <> 0 Then
        Rw = xx + 3
        Exit Sub
    End If
End If
'look third & fourth (fourth)-2
If yy = 3 Then
    If R(xx, yy) = 2 And R(xx + 1, yy + 1) = 2 And R(xx + 2, yy + 2) = 0 And R(xx + 3, yy + 3) = 0 Then
        Rw = xx + 3
        Exit Sub
    End If
End If
Next yy
Next xx
End Sub

Public Sub Look3Dia2()
For xx = 3 To 9
For yy = 0 To 3
'look first & second (first)
If R(xx, yy) = 0 And R(xx, yy + 1) <> 0 And R(xx - 1, yy + 1) = 0 And R(xx - 2, yy + 2) = 2 And R(xx - 3, yy + 3) = 2 Then
    Rw = xx
    Exit Sub
End If
'look first & second (second)
If R(xx, yy) = 0 And R(xx - 1, yy + 1) = 0 And R(xx - 1, yy + 2) <> 0 And R(xx - 2, yy + 2) = 2 And R(xx - 3, yy + 3) = 2 Then
    Rw = xx - 1
    Exit Sub
End If
'look first & third (first)
If R(xx, yy) = 0 And R(xx, yy + 1) <> 0 And R(xx - 1, yy + 1) = 2 And R(xx - 2, yy + 2) = 0 And R(xx - 3, yy + 3) = 2 Then
    Rw = xx
    Exit Sub
End If
'look first & third (third)
If R(xx, yy) = 0 And R(xx - 1, yy + 1) = 2 And R(xx - 2, yy + 2) = 0 And R(xx - 2, yy + 3) <> 0 And R(xx - 3, yy + 3) = 2 Then
    Rw = xx - 2
    Exit Sub
End If
'look first & fourth (first)
If R(xx, yy) = 0 And R(xx, yy + 1) <> 0 And R(xx - 1, yy + 1) = 2 And R(xx - 2, yy + 2) = 2 And R(xx - 3, yy + 3) = 0 Then
    Rw = xx
    Exit Sub
End If
'look first & fourth (fourth)
If yy < 3 Then
    If R(xx, yy) = 0 And R(xx - 1, yy + 1) = 2 And R(xx - 2, yy + 2) = 2 And R(xx - 3, yy + 3) = 0 And R(xx - 3, yy + 4) <> 0 Then
        Rw = xx - 3
        Exit Sub
    End If
End If
'look first & fourth (fourth)-2
If yy = 3 Then
    If R(xx, yy) = 0 And R(xx - 1, yy + 1) = 2 And R(xx - 2, yy + 2) = 2 And R(xx - 3, yy + 3) = 0 Then
        Rw = xx - 3
        Exit Sub
    End If
End If
'look second & third (second)
If R(xx, yy) = 2 And R(xx - 1, yy + 1) = 0 And R(xx - 1, yy + 2) <> 0 And R(xx - 2, yy + 2) = 0 And R(xx - 3, yy + 3) = 2 Then
    Rw = xx - 1
    Exit Sub
End If
'look second & third (third)
If R(xx, yy) = 2 And R(xx - 1, yy + 1) = 0 And R(xx - 2, yy + 2) = 0 And R(xx - 2, yy + 3) <> 0 And R(xx - 3, yy + 3) = 2 Then
    Rw = xx - 2
    Exit Sub
End If
'look second & fourth (second)
If R(xx, yy) = 2 And R(xx - 1, yy + 1) = 0 And R(xx - 1, yy + 2) <> 0 And R(xx - 2, yy + 2) = 2 And R(xx - 3, yy + 3) = 0 Then
    Rw = xx - 1
    Exit Sub
End If
'look second & fourth (fourth)
If yy < 3 Then
If R(xx, yy) = 2 And R(xx - 1, yy + 1) = 0 And R(xx - 2, yy + 2) = 2 And R(xx - 3, yy + 3) = 0 And R(xx - 3, yy + 4) <> 0 Then
    Rw = xx - 3
    Exit Sub
End If
End If
'look second & fourth (fourth)-2
If yy = 3 Then
If R(xx, yy) = 2 And R(xx - 1, yy + 1) = 0 And R(xx - 2, yy + 2) = 2 And R(xx - 3, yy + 3) = 0 Then
    Rw = xx - 3
    Exit Sub
End If
End If
'look third & fourth (third)
If R(xx, yy) = 2 And R(xx - 1, yy + 1) = 2 And R(xx - 2, yy + 2) = 0 And R(xx - 2, yy + 3) <> 0 And R(xx - 3, yy + 3) = 0 Then
    Rw = xx - 2
    Exit Sub
End If
'look third & fourth (fourth)
If yy < 3 Then
If R(xx, yy) = 2 And R(xx - 1, yy + 1) = 2 And R(xx - 2, yy + 2) = 0 And R(xx - 3, yy + 3) = 0 And R(xx - 3, yy + 4) <> 0 Then
    Rw = xx - 3
    Exit Sub
End If
End If
'look third & fourth (fourth)-2
If yy = 3 Then
If R(xx, yy) = 2 And R(xx - 1, yy + 1) = 2 And R(xx - 2, yy + 2) = 0 And R(xx - 3, yy + 3) = 0 Then
    Rw = xx - 3
    Exit Sub
End If
End If
Next yy
Next xx
End Sub

Public Sub LookPlayer3Hor()
'look if player can make 3 in a row horizontal
' if place below <> 0
For xx = 0 To 6
For yy = 0 To 6
If yy < 6 Then
'look first & second
If R(xx, yy) = 0 And R(xx, yy + 1) <> 0 And R(xx + 1, yy) = 0 And R(xx + 2, yy) = 1 And R(xx + 3, yy) = 1 Then
    Rw = xx
    Exit Sub
End If
'look first & second (2)
If R(xx, yy) = 0 And R(xx + 1, yy) = 0 And R(xx + 1, yy + 1) <> 0 And R(xx + 2, yy) = 1 And R(xx + 3, yy) = 1 Then
    Rw = xx + 1
    Exit Sub
End If
'look first & third
If R(xx, yy) = 0 And R(xx, yy + 1) <> 0 And R(xx + 1, yy) = 1 And R(xx + 2, yy) = 0 And R(xx + 3, yy) = 1 Then
    Rw = xx
    Exit Sub
End If
'look first & third (2)
If R(xx, yy) = 0 And R(xx + 1, yy) = 1 And R(xx + 2, yy) = 0 And R(xx + 2, yy + 1) <> 0 And R(xx + 3, yy) = 1 Then
    Rw = xx + 2
    Exit Sub
End If
'look first & fourth
If R(xx, yy) = 0 And R(xx, yy + 1) <> 0 And R(xx + 1, yy) = 1 And R(xx + 2, yy) = 1 And R(xx + 3, yy) = 0 Then
    Rw = xx
    Exit Sub
End If
'look first & fourth (2)
If R(xx, yy) = 0 And R(xx + 1, yy) = 1 And R(xx + 2, yy) = 1 And R(xx + 3, yy) = 0 And R(xx + 3, yy + 1) <> 0 Then
    Rw = xx + 3
    Exit Sub
End If
'look second & third
If R(xx, yy) = 1 And R(xx + 1, yy) = 0 And R(xx + 1, yy + 1) <> 0 And R(xx + 2, yy) = 0 And R(xx + 3, yy) = 1 Then
    Rw = xx + 1
    Exit Sub
End If
'look second & third (2)
If R(xx, yy) = 1 And R(xx + 1, yy) = 0 And R(xx + 2, yy) = 0 And R(xx + 2, yy + 1) <> 0 And R(xx + 3, yy) = 1 Then
    Rw = xx + 2
    Exit Sub
End If
'look second & fourth
If R(xx, yy) = 1 And R(xx + 1, yy) = 0 And R(xx + 1, yy + 1) <> 0 And R(xx + 2, yy) = 1 And R(xx + 3, yy) = 0 Then
    Rw = xx + 1
    Exit Sub
End If
'look second & fourth (2)
If R(xx, yy) = 1 And R(xx + 1, yy) = 0 And R(xx + 2, yy) = 1 And R(xx + 3, yy) = 0 And R(xx + 3, yy + 1) <> 0 Then
    Rw = xx + 3
    Exit Sub
End If
'look third & fourth
If R(xx, yy) = 1 And R(xx + 1, yy) = 1 And R(xx + 2, yy) = 0 And R(xx + 2, yy + 1) <> 0 And R(xx + 3, yy) = 0 Then
    Rw = xx + 2
    Exit Sub
End If
'look third & fourth (2)
If R(xx, yy) = 1 And R(xx + 1, yy) = 1 And R(xx + 2, yy) = 0 And R(xx + 3, yy) = 0 And R(xx + 3, yy + 1) <> 0 Then
    Rw = xx + 3
    Exit Sub
End If
End If

If yy = 6 Then
'look first & second
If R(xx, yy) = 0 And R(xx + 1, yy) = 0 And R(xx + 2, yy) = 1 And R(xx + 3, yy) = 1 Then
    Rw = xx
    Exit Sub
End If
'look first & third
If R(xx, yy) = 0 And R(xx + 1, yy) = 1 And R(xx + 2, yy) = 0 And R(xx + 3, yy) = 1 Then
    Rw = xx
    Exit Sub
End If
'look first & fourth
If R(xx, yy) = 0 And R(xx + 1, yy) = 1 And R(xx + 2, yy) = 1 And R(xx + 3, yy) = 0 Then
    Rw = xx
    Exit Sub
End If
'look second & third
If R(xx, yy) = 1 And R(xx + 1, yy) = 0 And R(xx + 2, yy) = 0 And R(xx + 3, yy) = 1 Then
    Rw = xx + 1
    Exit Sub
End If
'look second & fourth
If R(xx, yy) = 1 And R(xx + 1, yy) = 0 And R(xx + 2, yy) = 1 And R(xx + 3, yy) = 0 Then
    Rw = xx + 1
    Exit Sub
End If
'look third & fourth
If R(xx, yy) = 1 And R(xx + 1, yy) = 1 And R(xx + 2, yy) = 0 And R(xx + 3, yy) = 0 Then
    Rw = xx + 2
    Exit Sub
End If
End If
Next yy
Next xx
End Sub

Public Sub LookPlayer3Vert() 'look to make 3 in a row vertically
For xx = 0 To 9
For yy = 0 To 3
If R(xx, yy) = 0 And R(xx, yy + 1) = 0 And R(xx, yy + 2) = 1 And R(xx, yy + 3) = 1 Then
    Rw = xx
    Exit Sub
End If
Next yy
Next xx
End Sub

Public Sub LookPlayer3Dia1()
'look if player can make 3 in a row diagonal 1
For xx = 0 To 6
For yy = 0 To 3
'look first & second (first)
If R(xx, yy) = 0 And R(xx, yy + 1) <> 0 And R(xx + 1, yy + 1) = 0 And R(xx + 2, yy + 2) = 1 And R(xx + 3, yy + 3) = 1 Then
    Rw = xx
    Exit Sub
End If
'look first & second (second)
If R(xx, yy) = 0 And R(xx + 1, yy + 1) = 0 And R(xx + 1, yy + 2) <> 0 And R(xx + 2, yy + 2) = 1 And R(xx + 3, yy + 3) = 1 Then
    Rw = xx + 1
    Exit Sub
End If
'look first & third (first)
If R(xx, yy) = 0 And R(xx, yy + 1) <> 0 And R(xx + 1, yy + 1) = 1 And R(xx + 2, yy + 2) = 0 And R(xx + 3, yy + 3) = 1 Then
    Rw = xx
    Exit Sub
End If
'look first & third (third)
If R(xx, yy) = 0 And R(xx + 1, yy + 1) = 1 And R(xx + 2, yy + 2) = 0 And R(xx + 2, yy + 3) <> 0 And R(xx + 2, yy + 3) <> 0 And R(xx + 3, yy + 3) = 1 Then
    Rw = xx + 2
    Exit Sub
End If
'look first & fourth (first)
If R(xx, yy) = 0 And R(xx, yy + 1) <> 0 And R(xx + 1, yy + 1) = 1 And R(xx + 2, yy + 2) = 1 And R(xx + 3, yy + 3) = 0 Then
    Rw = xx
    Exit Sub
End If
'look first & fourth (fourth)
If yy < 3 Then
    If R(xx, yy) = 0 And R(xx + 1, yy + 1) = 1 And R(xx + 2, yy + 2) = 1 And R(xx + 3, yy + 3) = 0 And R(xx + 3, yy + 4) <> 0 Then
        Rw = xx + 3
        Exit Sub
    End If
End If
'look first & fourth (fourth)-2
If yy = 3 Then
    If R(xx, yy) = 0 And R(xx + 1, yy + 1) = 1 And R(xx + 2, yy + 2) = 1 And R(xx + 3, yy + 3) = 0 Then
        Rw = xx + 3
        Exit Sub
    End If
End If
'look second & third (second)
    If R(xx, yy) = 1 And R(xx + 1, yy + 1) = 0 And R(xx + 1, yy + 2) <> 0 And R(xx + 2, yy + 2) = 0 And R(xx + 3, yy + 3) = 1 Then
        Rw = xx + 1
        Exit Sub
    End If
'look second & third (third)
    If R(xx, yy) = 1 And R(xx + 1, yy + 1) = 0 And R(xx + 2, yy + 2) = 0 And R(xx + 2, yy + 3) <> 0 And R(xx + 3, yy + 3) = 1 Then
        Rw = xx + 2
        Exit Sub
    End If
'look second & fourth (second)
    If R(xx, yy) = 1 And R(xx + 1, yy + 1) = 0 And R(xx + 1, yy + 2) <> 0 And R(xx + 2, yy + 2) = 1 And R(xx + 3, yy + 3) = 0 Then
        Rw = xx + 1
        Exit Sub
    End If
'look second & fourth (fourth)
If yy < 3 Then
    If R(xx, yy) = 1 And R(xx + 1, yy + 1) = 0 And R(xx + 2, yy + 2) = 1 And R(xx + 3, yy + 3) = 0 And R(xx + 3, yy + 4) <> 0 Then
        Rw = xx + 3
        Exit Sub
    End If
End If
 'look second & fourth (fourth)-2
If yy = 3 Then
    If R(xx, yy) = 1 And R(xx + 1, yy + 1) = 0 And R(xx + 2, yy + 2) = 1 And R(xx + 3, yy + 3) = 0 Then
        Rw = xx + 3
        Exit Sub
    End If
End If
'look third & fourth (third)
    If R(xx, yy) = 1 And R(xx + 1, yy + 1) = 1 And R(xx + 2, yy + 2) = 0 And R(xx + 2, yy + 3) <> 0 And R(xx + 3, yy + 3) = 0 Then
        Rw = xx + 2
        Exit Sub
    End If
'look third & fourth (fourth)
If yy < 3 Then
    If R(xx, yy) = 1 And R(xx + 1, yy + 1) = 1 And R(xx + 2, yy + 2) = 0 And R(xx + 3, yy + 3) = 0 And R(xx + 3, yy + 4) <> 0 Then
        Rw = xx + 3
        Exit Sub
    End If
End If
'look third & fourth (fourth)-2
If yy = 3 Then
    If R(xx, yy) = 1 And R(xx + 1, yy + 1) = 1 And R(xx + 2, yy + 2) = 0 And R(xx + 3, yy + 3) = 0 Then
        Rw = xx + 3
        Exit Sub
    End If
End If
Next yy
Next xx
End Sub

Public Sub LookPlayer3Dia2()
For xx = 3 To 9
For yy = 0 To 3
'look first & second (first)
If R(xx, yy) = 0 And R(xx, yy + 1) <> 0 And R(xx - 1, yy + 1) = 0 And R(xx - 2, yy + 2) = 1 And R(xx - 3, yy + 3) = 1 Then
    Rw = xx
    Exit Sub
End If
'look first & second (second)
If R(xx, yy) = 0 And R(xx - 1, yy + 1) = 0 And R(xx - 1, yy + 2) <> 0 And R(xx - 2, yy + 2) = 1 And R(xx - 3, yy + 3) = 1 Then
    Rw = xx - 1
    Exit Sub
End If
'look first & third (first)
If R(xx, yy) = 0 And R(xx, yy + 1) <> 0 And R(xx - 1, yy + 1) = 1 And R(xx - 2, yy + 2) = 0 And R(xx - 3, yy + 3) = 1 Then
    Rw = xx
    Exit Sub
End If
'look first & third (third)
If R(xx, yy) = 0 And R(xx - 1, yy + 1) = 1 And R(xx - 2, yy + 2) = 0 And R(xx - 2, yy + 3) <> 0 And R(xx - 3, yy + 3) = 1 Then
    Rw = xx - 2
    Exit Sub
End If
'look first & fourth (first)
If R(xx, yy) = 0 And R(xx, yy + 1) <> 0 And R(xx - 1, yy + 1) = 1 And R(xx - 2, yy + 2) = 1 And R(xx - 3, yy + 3) = 0 Then
    Rw = xx
    Exit Sub
End If
'look first & fourth (fourth)
If yy < 3 Then
    If R(xx, yy) = 0 And R(xx - 1, yy + 1) = 1 And R(xx - 2, yy + 2) = 1 And R(xx - 3, yy + 3) = 0 And R(xx - 3, yy + 4) <> 0 Then
        Rw = xx - 3
        Exit Sub
    End If
End If
'look first & fourth (fourth)-2
If yy = 3 Then
    If R(xx, yy) = 0 And R(xx - 1, yy + 1) = 1 And R(xx - 2, yy + 2) = 1 And R(xx - 3, yy + 3) = 0 Then
        Rw = xx - 3
        Exit Sub
    End If
End If
'look second & third (second)
If R(xx, yy) = 1 And R(xx - 1, yy + 1) = 0 And R(xx - 1, yy + 2) <> 0 And R(xx - 2, yy + 2) = 0 And R(xx - 3, yy + 3) = 1 Then
    Rw = xx - 1
    Exit Sub
End If
'look second & third (third)
If R(xx, yy) = 1 And R(xx - 1, yy + 1) = 0 And R(xx - 2, yy + 2) = 0 And R(xx - 2, yy + 3) <> 0 And R(xx - 3, yy + 3) = 1 Then
    Rw = xx - 2
    Exit Sub
End If
'look second & fourth (second)
If R(xx, yy) = 1 And R(xx - 1, yy + 1) = 0 And R(xx - 1, yy + 2) <> 0 And R(xx - 2, yy + 2) = 1 And R(xx - 3, yy + 3) = 0 Then
    Rw = xx - 1
    Exit Sub
End If
'look second & fourth (fourth)
If yy < 3 Then
If R(xx, yy) = 1 And R(xx - 1, yy + 1) = 0 And R(xx - 2, yy + 2) = 1 And R(xx - 3, yy + 3) = 0 And R(xx - 3, yy + 4) <> 0 Then
    Rw = xx - 3
    Exit Sub
End If
End If
'look second & fourth (fourth)-2
If yy = 3 Then
If R(xx, yy) = 1 And R(xx - 1, yy + 1) = 0 And R(xx - 2, yy + 2) = 1 And R(xx - 3, yy + 3) = 0 Then
    Rw = xx - 3
    Exit Sub
End If
End If
'look third & fourth (third)
If R(xx, yy) = 1 And R(xx - 1, yy + 1) = 1 And R(xx - 2, yy + 2) = 0 And R(xx - 2, yy + 3) <> 0 And R(xx - 3, yy + 3) = 0 Then
    Rw = xx - 2
    Exit Sub
End If
'look third & fourth (fourth)
If yy < 3 Then
If R(xx, yy) = 1 And R(xx - 1, yy + 1) = 1 And R(xx - 2, yy + 2) = 0 And R(xx - 3, yy + 3) = 0 And R(xx - 3, yy + 4) <> 0 Then
    Rw = xx - 3
    Exit Sub
End If
End If
'look third & fourth (fourth)-2
If yy = 3 Then
If R(xx, yy) = 1 And R(xx - 1, yy + 1) = 1 And R(xx - 2, yy + 2) = 0 And R(xx - 3, yy + 3) = 0 Then
    Rw = xx - 3
    Exit Sub
End If
End If
Next yy
Next xx
End Sub

Public Sub BlockPlayerHor()
For xx = 0 To 8
For yy = 0 To 6
If yy < 6 Then
    If R(xx, yy) = 1 And R(xx + 1, yy) = 0 And R(xx + 1, yy) <> 0 Then
    Rw = xx + 1
    Exit Sub
    End If
End If
If yy = 6 Then
    If R(xx, yy) = 1 And R(xx + 1, yy) = 0 Then
    Rw = xx + 1
    Exit Sub
    End If
End If
Next yy
Next xx
End Sub

Public Sub BlockPlayerVert()
For xx = 0 To 9
For yy = 0 To 6
    If R(xx, yy) = 1 Then
    Rw = xx
    Exit Sub
    End If
Next yy
Next xx
End Sub
