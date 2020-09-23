VERSION 5.00
Begin VB.Form Bases 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Super Calculatrice 7.0"
   ClientHeight    =   3870
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   5640
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   3870
   ScaleWidth      =   5640
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Décimal-base"
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Base-décimal"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   2040
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   3360
      Width           =   3735
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   2880
      Width           =   3735
   End
   Begin VB.Line Line7 
      X1              =   3120
      X2              =   3120
      Y1              =   120
      Y2              =   1440
   End
   Begin VB.Line Line6 
      X1              =   5400
      X2              =   5400
      Y1              =   120
      Y2              =   1440
   End
   Begin VB.Line Line5 
      X1              =   120
      X2              =   5400
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   120
      Y1              =   1440
      Y2              =   120
   End
   Begin VB.Line Line3 
      X1              =   5400
      X2              =   120
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   5400
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line1 
      X1              =   1320
      X2              =   1320
      Y1              =   120
      Y2              =   1440
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Nombre Décimal"
      Height          =   255
      Left            =   1560
      TabIndex        =   12
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Base"
      Height          =   255
      Left            =   4080
      TabIndex        =   11
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Nombre"
      Height          =   255
      Left            =   1920
      TabIndex        =   10
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Base de ce nombre"
      Height          =   375
      Left            =   3600
      TabIndex        =   9
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Décimal-base"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Base-Décimal"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Réponse"
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   2520
      Width           =   3735
   End
   Begin VB.Menu prog 
      Caption         =   "programme"
      Begin VB.Menu quit 
         Caption         =   "Quitter"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "Bases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text3.Text = ConvertToBase(Val(Text1.Text), Val(Text2.Text))
End Sub

Private Sub Command2_Click()
Text3.Text = ConvertFromBase(Val(Text1.Text), Val(Text2.Text))
End Sub

Private Sub Command3_Click()
quit_Click
End Sub

Private Sub Form_Initialize()
Call InitCommonControls
End Sub
Private Sub quit_Click()
Bases.Hide
calc7.Show
End Sub
