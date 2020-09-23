VERSION 5.00
Begin VB.Form StartForm 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4035
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   ScaleHeight     =   269
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   315
      ScaleHeight     =   195
      ScaleWidth      =   240
      TabIndex        =   9
      Top             =   1035
      Width           =   240
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   225
      Top             =   315
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "Impossible to win !"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   300
      Index           =   4
      Left            =   2655
      TabIndex        =   4
      Top             =   2700
      Width           =   2550
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "Sweat and tears..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   300
      Index           =   3
      Left            =   2655
      TabIndex        =   3
      Top             =   2385
      Width           =   2550
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "It can be done !"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   300
      Index           =   2
      Left            =   2655
      TabIndex        =   2
      Top             =   2070
      Width           =   2550
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "You win, no doubt..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   300
      Index           =   1
      Left            =   2655
      TabIndex        =   1
      Top             =   1755
      Width           =   2550
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "Impossible to loose !"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   300
      Index           =   0
      Left            =   2655
      TabIndex        =   0
      Top             =   1440
      Width           =   2550
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Start game"
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
      Left            =   2925
      TabIndex        =   8
      Top             =   3240
      Width           =   2130
   End
   Begin VB.Image Image4 
      Height          =   3240
      Left            =   6075
      Picture         =   "StartForm.frx":0000
      Top             =   450
      Width           =   1500
   End
   Begin VB.Image Image3 
      Height          =   3240
      Left            =   495
      Picture         =   "StartForm.frx":1E22
      Top             =   450
      Width           =   1500
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select difficulty"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   420
      Left            =   2475
      TabIndex        =   7
      Top             =   1035
      Width           =   3030
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Coded by Stephan Swertvaegher"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   375
      Index           =   1
      Left            =   1620
      TabIndex        =   6
      Top             =   315
      Width           =   4740
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Coded by Stephan Swertvaegher"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Index           =   0
      Left            =   1665
      TabIndex        =   5
      Top             =   330
      Width           =   4740
   End
   Begin VB.Image Image2 
      Height          =   2085
      Left            =   2385
      Top             =   990
      Width           =   3210
   End
End
Attribute VB_Name = "StartForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Option1(Difficulty).Value = True
Picture1.SetFocus
End Sub

Private Sub Form_Load()
T3D StartForm, Image2, 3, T3dRaiseInset, T3dF1
T = 0
End Sub

Private Sub Label3_Click()
For xx = 0 To 4
If Option1(xx).Value = True Then
Difficulty = xx
End If
Next xx
StartForm.Hide
RowFrm.Label1(0).Caption = "Difficulty: " & Str(Difficulty)
RowFrm.Label1(1).Caption = Dif(Difficulty)
RowFrm.Timer5.Enabled = True
RowFrm.Command2.Enabled = False
End Sub

Private Sub Option1_Click(Index As Integer)
Picture1.SetFocus
End Sub

Private Sub Timer1_Timer()
Label3.ForeColor = Col(T)
T = T + 1
If T = 48 Then T = 0
End Sub
