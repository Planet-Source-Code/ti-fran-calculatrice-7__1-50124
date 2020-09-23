VERSION 5.00
Begin VB.Form sys 
   Caption         =   "Super Calculatrice 7.0"
   ClientHeight    =   2985
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   4710
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2985
   ScaleWidth      =   4710
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Quitter"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   2520
      Width           =   4455
   End
   Begin VB.Frame Frame4 
      Caption         =   "Systèmes à 2 équations"
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.TextBox Text13 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2640
         TabIndex        =   9
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox Text12 
         Enabled         =   0   'False
         Height          =   375
         Left            =   720
         TabIndex        =   8
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox Text11 
         Height          =   375
         Left            =   3240
         TabIndex        =   7
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   1920
         TabIndex        =   6
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox Text9 
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   3240
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   1920
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "résoudre"
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   1200
         Width           =   3855
      End
      Begin VB.Label Label8 
         Caption         =   "y ="
         Height          =   375
         Left            =   2400
         TabIndex        =   15
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label Label7 
         Caption         =   " x ="
         Height          =   375
         Left            =   360
         TabIndex        =   14
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   " y ="
         Height          =   375
         Left            =   2880
         TabIndex        =   13
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "  x +"
         Height          =   375
         Left            =   1560
         TabIndex        =   12
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   " y ="
         Height          =   375
         Left            =   2880
         TabIndex        =   11
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "  x +"
         Height          =   375
         Index           =   0
         Left            =   1560
         TabIndex        =   10
         Top             =   360
         Width           =   375
      End
   End
End
Attribute VB_Name = "sys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
sys.Hide
calc7.Show
End Sub

Private Sub Command3_Click()
Dim a, b, c, d, e, f, x, y As Double
a = Val(Text6)
b = Val(Text7)
c = Val(Text8)
d = Val(Text9)
e = Val(Text10)
f = Val(Text11)
x = (e * c - b * f) / (a * e - d * b)
y = (f * a - c * d) / (a * e - d * b)
Text12 = x
Text13 = y
End Sub
Private Sub Form_Initialize()
Call InitCommonControls
End Sub
