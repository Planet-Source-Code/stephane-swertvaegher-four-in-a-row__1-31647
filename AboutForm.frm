VERSION 5.00
Begin VB.Form AboutForm 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4260
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   ScaleHeight     =   284
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   421
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image3 
      Height          =   1335
      Left            =   315
      Picture         =   "AboutForm.frx":0000
      ToolTipText     =   "About..."
      Top             =   180
      Width           =   990
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0080FFFF&
      Height          =   2355
      Left            =   675
      TabIndex        =   1
      Top             =   1665
      Width           =   4875
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "4 in a row"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   465
      Left            =   2115
      TabIndex        =   0
      Top             =   495
      Width           =   3075
   End
End
Attribute VB_Name = "AboutForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Image3_Click()
AboutForm.Hide
End Sub
