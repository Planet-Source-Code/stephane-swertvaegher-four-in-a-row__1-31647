VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form RowFrm 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10830
   ForeColor       =   &H00000040&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   514
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   722
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command4 
      BackColor       =   &H0000C000&
      Caption         =   "About..."
      Height          =   555
      Left            =   8910
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6660
      Width           =   915
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000C0C0&
      Caption         =   "Change difficulty"
      Height          =   330
      Left            =   8550
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4815
      Width           =   1590
   End
   Begin VB.Timer Timer5 
      Interval        =   10
      Left            =   2115
      Top             =   810
   End
   Begin VB.Timer Timer4 
      Interval        =   1
      Left            =   1665
      Top             =   810
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   6000
      Left            =   270
      ScaleHeight     =   400
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   496
      TabIndex        =   5
      Top             =   1440
      Width           =   7440
      Begin VB.Image Image1 
         Height          =   735
         Index           =   0
         Left            =   0
         Top             =   0
         Width           =   735
      End
      Begin VB.Image Image2 
         Height          =   735
         Index           =   0
         Left            =   0
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000C0C0&
      Caption         =   "New Game"
      Height          =   330
      Left            =   8550
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5265
      Width           =   1590
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   1215
      Top             =   810
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   765
      Top             =   810
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   225
      Top             =   810
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5130
      Top             =   3555
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   49
      ImageHeight     =   49
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "4inrow.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "4inrow.frx":057C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "4inrow.frx":0AC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "4inrow.frx":1050
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "4inrow.frx":16C0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "Exit Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8550
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5715
      Width           =   1590
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   420
      Left            =   7920
      TabIndex        =   7
      Top             =   3915
      Width           =   2670
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   330
      Left            =   7920
      TabIndex        =   6
      Top             =   2565
      Width           =   2670
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   600
      Left            =   7920
      TabIndex        =   3
      Top             =   3060
      Width           =   2670
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   1
      Left            =   7920
      TabIndex        =   2
      Top             =   1845
      Width           =   2760
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Difficulty:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   330
      Index           =   0
      Left            =   7965
      TabIndex        =   1
      Top             =   1530
      Width           =   2625
   End
   Begin VB.Image Title 
      Height          =   1110
      Left            =   1215
      Picture         =   "4inrow.frx":1D18
      Top             =   270
      Width           =   8175
   End
End
Attribute VB_Name = "RowFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim temp$
Msbox "Are you sure to quit" & vbCr & "4 in a row ?", "4 in a row", mbYesNo, mbQuestion
If MBReturn = 0 Then
End
End If
End Sub

Private Sub Command2_Click()
Timer4.Enabled = True
End Sub

Private Sub Command3_Click()
If PT = 2 Then
Command2_Click
Exit Sub
End If
    StartForm.Show 1
End Sub

Private Sub Command4_Click()
AboutForm.Show 1
End Sub

Private Sub Form_Load()
SetVar
ScreenSetUp
RowFrm.Show
Command2.Enabled = False
StartForm.Show 1
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.Caption = ""
End Sub

Private Sub Image1_Click(Index As Integer)
If PT <> 0 Then Exit Sub
If Timer1.Enabled = True Then Exit Sub
Rw = Index - (Int(Index / 10) * 10)
PlayerBall
End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If PT <> 0 Or Timer1.Enabled = True Then Label3.Caption = "": Exit Sub
Label3.Caption = "Row: " & (Index - (Int(Index / 10) * 10)) + 1
End Sub

Private Sub Image2_Click(Index As Integer)
If PT <> 0 Then Exit Sub
If Timer1.Enabled = True Then Exit Sub
Rw = Index
PlayerBall
End Sub

Private Sub Image2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If PT <> 0 Or Timer1.Enabled = True Then Label3.Caption = "": Exit Sub
Label3.Caption = "Row: " & (Index + 1)
End Sub

Private Sub PlayerBall()
If R(Rw, 0) <> 0 Then
Label2.Caption = "That's not allowed !"
Beep
Exit Sub
End If
Fall = 0
Turns = Turns + 1
Label4.Caption = "Turns: " & Turns
Label2.Caption = "You put a coin in" & vbCr & "row " & Rw + 1
Image1(Rw + (Fall * 10)).Picture = ImageList1.ListImages(2).Picture
If R(Rw, 1) = 0 Then
Timer1.Enabled = True
Else
R(Rw, 0) = 1
Control (1)
End If
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.Caption = ""
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.Caption = ""
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.Caption = ""
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.Caption = ""
End Sub

Private Sub Timer1_Timer()
Fall = Fall + 1
If Fall = 6 Then
    R(Rw, Fall) = 1
    Timer1.Enabled = False
    Image1((Fall * 10) + Rw).Picture = ImageList1.ListImages(2).Picture
    Image1(((Fall - 1) * 10) + Rw).Picture = ImageList1.ListImages(1).Picture
    Control (1)
        If PT = 2 Then
        Exit Sub
        End If
    PT = 1 'computers turn
    Label2.Caption = "My turn"
    DoEvents
    Fall = 0
    Timer2.Enabled = True
    Exit Sub
End If

If R(Rw, Fall + 1) <> 0 Then
    R(Rw, Fall) = 1
    Timer1.Enabled = False
    Image1((Fall * 10) + Rw).Picture = ImageList1.ListImages(2).Picture
    Image1(((Fall - 1) * 10) + Rw).Picture = ImageList1.ListImages(1).Picture
    Control (1)
        If PT = 2 Then
        Exit Sub
        End If
    PT = 1 'computers turn
    Label2.Caption = "My turn"
    DoEvents
    Fall = 0
    Timer2.Enabled = True
    Exit Sub
End If

    Image1((Fall * 10) + Rw).Picture = ImageList1.ListImages(2).Picture
    Image1(((Fall - 1) * 10) + Rw).Picture = ImageList1.ListImages(1).Picture

End Sub

Private Sub Timer2_Timer()
'Computer looks for row to put coin in
Rw = 255 'set rw = impossible
Select Case Difficulty
Case 0
    Easy
Case 1
    NotSoEasy
Case 2
    Hard
Case 3
    VeryHard
Case 4
    ImpossibleToWin
End Select
' nothing found, put coin random
If Rw = 255 Or Rw < 0 Then
DoAgain:
Randomize
Rw = Int(Rnd * 10)
    If R(Rw, 0) <> 0 Then GoTo DoAgain
End If
Timer2.Enabled = False
Fall = 0
Label2.Caption = "I put a coin in" & vbCr & "row " & Rw + 1
Image1(Rw + (Fall * 10)).Picture = ImageList1.ListImages(3).Picture
If R(Rw, 1) = 0 Then
Timer3.Enabled = True
Else
R(Rw, 0) = 2
    Timer3.Enabled = False
    Control (2)
        If PT = 2 Then
        Exit Sub
        End If
    PT = 0 'players turn
    Label2.Caption = "Your turn"
    DoEvents
    Fall = 0
End If
End Sub

Private Sub Timer3_Timer()
Fall = Fall + 1

If Fall = 6 Then
    R(Rw, Fall) = 2
    Image1((Fall * 10) + Rw).Picture = ImageList1.ListImages(3).Picture
    Image1(((Fall - 1) * 10) + Rw).Picture = ImageList1.ListImages(1).Picture
    Control (2)
    Timer3.Enabled = False
        If PT = 2 Then
        Exit Sub
        End If
    PT = 0 'players turn
    Label2.Caption = "Your turn"
    DoEvents
    Fall = 0
    Exit Sub
End If

If R(Rw, Fall + 1) <> 0 Then
    R(Rw, Fall) = 2
    Image1((Fall * 10) + Rw).Picture = ImageList1.ListImages(3).Picture
    Image1(((Fall - 1) * 10) + Rw).Picture = ImageList1.ListImages(1).Picture
    Control (2)
    Timer3.Enabled = False
        If PT = 2 Then
        Exit Sub
        End If
    PT = 0 'players turn
    Label2.Caption = "Your turn"
    DoEvents
    Fall = 0
    Exit Sub
End If

    Image1((Fall * 10) + Rw).Picture = ImageList1.ListImages(3).Picture
    Image1(((Fall - 1) * 10) + Rw).Picture = ImageList1.ListImages(1).Picture

End Sub

Private Sub Timer4_Timer()
Picture1.Top = Picture1.Top + 30
If Picture1.Top > RowFrm.ScaleHeight Then
    Timer4.Enabled = False
    SetVar
    StartForm.Show 1
    For xx = 0 To 69
    Image1(xx).Picture = ImageList1.ListImages(1).Picture
    Next xx
    Timer5.Enabled = True
End If
End Sub

Private Sub Timer5_Timer()
Picture1.Top = Picture1.Top - 30
If Picture1.Top <= 96 Then
Picture1.Top = 96
Timer5.Enabled = False
Label2.Caption = "Player, you start..."
Command2.Enabled = False
End If
End Sub

Private Sub Title_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.Caption = ""
End Sub
