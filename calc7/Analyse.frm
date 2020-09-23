VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form Analyse 
   Caption         =   "Super Calculatrice 7.0"
   ClientHeight    =   5070
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   6585
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   ScaleHeight     =   5070
   ScaleWidth      =   6585
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command8 
      Caption         =   "&Quitter"
      Height          =   375
      Left            =   3360
      TabIndex        =   8
      Top             =   1560
      Width           =   3135
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   2400
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Conversions entre bases 1 - 36"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   3135
   End
   Begin VB.CommandButton Command6 
      Caption         =   "C&ombinaisons"
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   1080
      Width           =   3135
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Résolution d'équation du second degré"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   3135
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Système d'équations"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   600
      Width           =   3135
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Zéro d'une fonction du premier degré"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Diviseurs d'un nombre"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Facteurs premiers d'un nombre"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "Écrivez-moi à franco44444@hotmail.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   360
      TabIndex        =   9
      Top             =   2880
      Width           =   5775
   End
   Begin VB.Shape Shape1 
      Height          =   2535
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   2400
      Width           =   6135
   End
   Begin VB.Menu prog 
      Caption         =   "programme"
      Begin VB.Menu quit 
         Caption         =   "Quitter"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "Analyse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Initialize()
Call InitCommonControls
End Sub
Private Sub Command1_Click()
factprem.Show
Analyse.Hide
End Sub

Private Sub Command2_Click()
Analyse.Hide
divnombre.Show
End Sub

Private Sub Command3_Click()
Analyse.Hide
zero.Show
End Sub

Private Sub Command4_Click()
Analyse.Hide
sys.Show
End Sub

Private Sub Command5_Click()
Analyse.Hide
frmQuadratic.Show
End Sub

Private Sub Command6_Click()
Analyse.Hide
cg.Show
End Sub

Private Sub Command7_Click()
Analyse.Hide
Bases.Show
End Sub

Private Sub Command8_Click()
Analyse.Hide
calc7.Show
End Sub

Private Sub quit_Click()
Analyse.Hide
calc7.Show
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
ProgressBar1.Value = ProgressBar1.Value + 1
If ProgressBar1.Value = 100 Then ProgressBar1.Value = 0
End Sub

Private Sub Timer2_Timer()
If Line1.X1 = 50 Then Line1.X1 = 0
Line1.X1 = Line1.X1 + 1
End Sub
