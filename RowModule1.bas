Attribute VB_Name = "Module1"
Public xx%, yy%, R(9, 6), Rw%, PT%, Fall%, Turns%
Public Difficulty%, Dif$(4)
Public Enum T3dFill
T3dF0
T3dF1
End Enum

Public Enum Borderstyle
T3dRaiseRaise
T3dRaiseInset
T3dInsetRaise
T3dInsetInset
T3dNone
End Enum

Public T%, Col&(47)

Public Sub ScreenSetUp()
Col(0) = 0: Col(1) = &H1010&: Col(2) = &H2020&
Col(3) = &H3030&: Col(4) = &H4040&: Col(5) = &H5050&
Col(6) = &H6060&: Col(7) = &H7070&: Col(8) = &H8080&
Col(9) = &H9090&: Col(10) = &HA0A0&: Col(11) = &HB0B0&
Col(12) = &HC0C0&: Col(13) = &HB0B0&: Col(14) = &HA0A0&
Col(15) = &H9090&: Col(16) = &H8080&: Col(17) = &H7070&
Col(18) = &H6060&: Col(19) = &H5050&: Col(20) = &H4040&
Col(21) = &H3030&: Col(22) = &H2020&: Col(23) = &H1010&
Col(24) = 0: Col(25) = &H101000: Col(26) = &H202000
Col(27) = &H303000: Col(28) = &H404000: Col(29) = &H505000
Col(30) = &H606000: Col(31) = &H707000: Col(32) = &H808000
Col(33) = &H909000: Col(34) = &HA0A000: Col(35) = &HB0B000
Col(36) = &HC0C000: Col(37) = &HB0B000: Col(38) = &HA0A000
Col(39) = &H909000: Col(40) = &H808000: Col(41) = &H707000
Col(42) = &H606000: Col(43) = &H505000: Col(44) = &H404000
Col(45) = &H303000: Col(46) = &H202000: Col(47) = &H101000
With RowFrm
.Move (Screen.Width / 2) - (.Width / 2), 120, 10830, 7635
For xx = 1 To 9
Load .Image2(xx)
.Image2(xx).Visible = True
Next xx
For xx = 1 To 69
Load .Image1(xx)
.Image1(xx).Visible = True
Next xx
.Image1(0).Move 0, 0
For xx = 0 To 9
For yy = 0 To 6
.Image1((yy * 10) + xx).Move (xx * 49), 49 + (yy * 49)
.Image1((yy * 10) + xx).Picture = .ImageList1.ListImages(1).Picture
Next yy
Next xx
For xx = 0 To 9
.Image2(xx).Move 0 + (xx * 49), 0
Next xx
ColBox RowFrm, 5, 90, 520, 502, 7, 50, 50, 0, 255, 128, 0
ColBox RowFrm, 520, 90, 715, 160, 7, 50, 50, 0, 255, 128, 0
ColBox RowFrm, 520, 160, 715, 502, 7, 50, 50, 0, 255, 128, 0
ColBox RowFrm, 0, 0, RowFrm.ScaleWidth - 1, RowFrm.ScaleHeight - 1, 7, 50, 50, 0, 64, 128, 50
ColBox StartForm, 0, 0, StartForm.ScaleWidth - 1, StartForm.ScaleHeight - 1, 7, 0, 50, 50, 0, 255, 128
ColBox AboutForm, 0, 0, AboutForm.ScaleWidth - 1, AboutForm.ScaleHeight - 1, 7, 0, 50, 50, 0, 255, 128
.Picture1.Top = RowFrm.ScaleHeight
End With
Difficulty = 2
With AboutForm
.Label1.Caption = "4 IN A ROW"
.Label2.Caption = "About..." & vbCr & vbCr & "4 in a row was a pupular game in the eighties. "
.Label2.Caption = .Label2.Caption & "Lots of people played it, just for fun, or in competition. "
.Label2.Caption = .Label2.Caption & "The main goal is to make 4 equal coins in a row, horizontal, vertical or diagonal. "
.Label2.Caption = .Label2.Caption & "Just wach out, 'cause the computer tries the same thing..." & vbCr
.Label2.Caption = .Label2.Caption & "Just click in a row, and a coin will appear. There are 5 levels in this game, starting from "
.Label2.Caption = .Label2.Caption & " very easy to very hard. Normally you can't win in level 5, but I won already... "
.Label2.Caption = .Label2.Caption & "Maybe the computer isn't so smart after all..." & vbCr & vbCr
.Label2.Caption = .Label2.Caption & "Now hit my face to continue..."

End With
End Sub

Public Sub SetVar()
With RowFrm
For xx = 0 To 9
For yy = 0 To 6
R(xx, yy) = 0
Next yy
Next xx
Dif(0) = "Impossible to loose !"
Dif(1) = "You win, no doubt..."
Dif(2) = "It can be done !"
Dif(3) = "Sweat and tears..."
Dif(4) = "Impossible to win !"
PT = 0 'players turn
.Timer1.Enabled = False
.Timer2.Enabled = False
.Timer3.Enabled = False
.Timer4.Enabled = False
.Timer5.Enabled = False
.Label1(0).Caption = "Difficulty: " & Difficulty
.Label3.Caption = ""
.Label4.Caption = ""
Turns = 0
End With
End Sub
Public Sub ColBar(Obj As Object, St%, H%, R%, g%, b%, RE%, GE%, BE%)
Dim H2%, H3%, IvR%, IvG%, IvB%
Obj.AutoRedraw = True
Obj.ScaleMode = 3 'pixel
H3 = Int(H / 2)
IvR = Int(RE - R) / H3
IvG = Int(GE - g) / H3
IvB = Int(BE - b) / H3
Do While H >= H3
Obj.Line (0, St + H2)-(Obj.ScaleWidth, St + H2), RGB(R, g, b)
Obj.Line (0, St + H)-(Obj.ScaleWidth, St + H), RGB(R, g, b)
H = H - 1
H2 = H2 + 1
R = R + IvR
g = g + IvG
b = b + IvB
Loop
End Sub

Public Sub ColBox(Obj As Object, BX%, BY%, EX%, EY%, H%, R%, g%, b%, RE%, GE%, BE%)
Dim H2%, H3%, IvR%, IvG%, IvB%
Obj.AutoRedraw = True
Obj.ScaleMode = 3 'pixel
H3 = Int(H / 2)
IvR = Int(RE - R) / H3
IvG = Int(GE - g) / H3
IvB = Int(BE - b) / H3
Do While H >= H3
Obj.Line (BX + H2, BY + H2)-(EX - H2, EY - H2), RGB(R, g, b), B
Obj.Line (BX + H, BY + H)-(EX - H, EY - H), RGB(R, g, b), B
H = H - 1
H2 = H2 + 1
R = R + IvR
g = g + IvG
b = b + IvB
Loop
End Sub

Public Sub Control(Pl%)
Dim x1%, x2%, x3%, x0%, y0%, y1%, y2%, y3%
'Control horizontal
For xx = 0 To 6
For yy = 0 To 6
    If R(xx, yy) = Pl And R(xx + 1, yy) = Pl And R(xx + 2, yy) = Pl And R(xx + 3, yy) = Pl Then
        x0 = xx: x1 = xx + 1: x2 = xx + 2: x3 = xx + 3
        y0 = yy: y1 = yy: y2 = yy: y3 = yy
        GoTo EndControll
    End If
Next yy
Next xx

'Control vertical
For xx = 0 To 9
For yy = 0 To 3
    If R(xx, yy) = Pl And R(xx, yy + 1) = Pl And R(xx, yy + 2) = Pl And R(xx, yy + 3) = Pl Then
        x0 = xx: x1 = xx: x2 = xx: x3 = xx
        y0 = yy: y1 = yy + 1: y2 = yy + 2: y3 = yy + 3
        GoTo EndControll
    End If
Next yy
Next xx

'Control diagonal 1
For xx = 0 To 6
For yy = 0 To 3
    If R(xx, yy) = Pl And R(xx + 1, yy + 1) = Pl And R(xx + 2, yy + 2) = Pl And R(xx + 3, yy + 3) = Pl Then
        x0 = xx: x1 = xx + 1: x2 = xx + 2: x3 = xx + 3
        y0 = yy: y1 = yy + 1: y2 = yy + 2: y3 = yy + 3
        GoTo EndControll
    End If
Next yy
Next xx

'Control diagonal 2
For xx = 3 To 9
For yy = 0 To 3
    If R(xx, yy) = Pl And R(xx - 1, yy + 1) = Pl And R(xx - 2, yy + 2) = Pl And R(xx - 3, yy + 3) = Pl Then
        x0 = xx: x1 = xx - 1: x2 = xx - 2: x3 = xx - 3
        y0 = yy: y1 = yy + 1: y2 = yy + 2: y3 = yy + 3
    GoTo EndControll
    End If
Next yy
Next xx

Exit Sub
EndControll:
With RowFrm
If Pl = 1 Then
.Label2.Caption = "Player, you have" & vbCr & "4 IN A ROW"
Else
.Label2.Caption = "Hey man... I have" & vbCr & "4 IN A ROW"
End If
.Image1((y0 * 10) + x0).Picture = .ImageList1.ListImages(3 + Pl).Picture
.Image1((y1 * 10) + x1).Picture = .ImageList1.ListImages(3 + Pl).Picture
.Image1((y2 * 10) + x2).Picture = .ImageList1.ListImages(3 + Pl).Picture
.Image1((y3 * 10) + x3).Picture = .ImageList1.ListImages(3 + Pl).Picture
.Command2.Enabled = True
End With
PT = 2 'player & computer cannot play anymore
End Sub


'My award winning T3D-function, a bit modified
' for a black background...
Public Function T3D(Obj0 As Object, Obj As Object, Bev%, Optional Style3D As Borderstyle, Optional T3dFilled As T3dFill)
Dim R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, R4%, G4%, B4%
Dim T3Dxx%
On Error Resume Next

Obj.Borderstyle = 0 'no border

If IsMissing(Style3D) Then Style3D = 0

If Style3D > 4 Then Style3D = 3

If Style3D = 0 Then 'RaiseRaise
R1 = 164: R2 = 96: R3 = 164: R4 = 96
End If
If Style3D = 1 Then 'RaiseInset
R1 = 164: R2 = 96: R4 = 164: R3 = 96
End If
If Style3D = 2 Then 'InsetRaise
R2 = 164: R1 = 96: R3 = 164: R4 = 96
End If
If Style3D = 3 Then 'InsetInset
R2 = 164: R1 = 96: R4 = 164: R3 = 96
End If
If Style3D = 4 Then 'No Border
R1 = 192: R2 = 192: R3 = 192: R4 = 192
End If

G1 = R1: B1 = R1
G2 = R2: B2 = R2
G3 = R3: B3 = R3
G4 = R4: B4 = R4
Bev = Bev + 1
T3Dxx = Bev
'Outer
If IsMissing(T3dFilled) Or T3dFilled = 0 Then
    Obj0.Line (Obj.Left - Bev, Obj.Top - Bev)-(Obj.Left - Bev, Obj.Top + Obj.Height + Bev), RGB(R1, G1, B1)
    Obj0.Line (Obj.Left - Bev, Obj.Top - Bev)-(Obj.Left + Obj.Width + Bev, Obj.Top - Bev), RGB(R1, G1, B1)
    Obj0.Line (Obj.Left + Obj.Width + Bev, Obj.Top - Bev)-(Obj.Left + Obj.Width + Bev, Obj.Top + Obj.Height + Bev), RGB(R2, G2, B2)
    Obj0.Line (Obj.Left - Bev, Obj.Top + Obj.Height + Bev)-(Obj.Left + Obj.Width + Bev + 1, Obj.Top + Obj.Height + Bev), RGB(R2, G2, B2)
Else
For Bev = T3Dxx To 1 Step -1
    Obj0.Line (Obj.Left - Bev, Obj.Top - Bev)-(Obj.Left - Bev, Obj.Top + Obj.Height + Bev), RGB(R1, G1, B1)
    Obj0.Line (Obj.Left - Bev, Obj.Top - Bev)-(Obj.Left + Obj.Width + Bev, Obj.Top - Bev), RGB(R1, G1, B1)
    Obj0.Line (Obj.Left + Obj.Width + Bev, Obj.Top - Bev)-(Obj.Left + Obj.Width + Bev, Obj.Top + Obj.Height + Bev), RGB(R2, G2, B2)
    Obj0.Line (Obj.Left - Bev, Obj.Top + Obj.Height + Bev)-(Obj.Left + Obj.Width + Bev + 1, Obj.Top + Obj.Height + Bev), RGB(R2, G2, B2)
Next Bev
End If
'Inner
    Obj0.Line (Obj.Left - 1, Obj.Top - 1)-(Obj.Left - 1, Obj.Top + Obj.Height + 1), RGB(R3, G3, B3)
    Obj0.Line (Obj.Left - 1, Obj.Top - 1)-(Obj.Left + Obj.Width + 1, Obj.Top - 1), RGB(R3, G3, B3)
    Obj0.Line (Obj.Left + Obj.Width + 1, Obj.Top - 1)-(Obj.Left + Obj.Width + 1, Obj.Top + Obj.Height + 1), RGB(R4, G4, B4)
    Obj0.Line (Obj.Left - 1, Obj.Top + Obj.Height + 1)-(Obj.Left + Obj.Width + 2, Obj.Top + Obj.Height + 1), RGB(R4, G4, B4)
End Function
