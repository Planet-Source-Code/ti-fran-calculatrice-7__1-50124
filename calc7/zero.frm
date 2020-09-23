VERSION 5.00
Begin VB.Form zero 
   Caption         =   "Super Calculatrice 7.0"
   ClientHeight    =   2385
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2385
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Zéro d'une fonction"
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   4455
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   480
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2280
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Calculer"
         Height          =   375
         Left            =   480
         TabIndex        =   3
         Top             =   840
         Width           =   3135
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   375
         Left            =   480
         TabIndex        =   2
         Top             =   1320
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "0 ="
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "  x +"
         Height          =   375
         Left            =   1920
         TabIndex        =   6
         Top             =   480
         Width           =   375
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Quitter"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   4455
   End
End
Attribute VB_Name = "zero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
zero.Hide
calc7.Show
End Sub

Private Sub Command2_Click()
Text3.Text = -(Val(Text2.Text)) / Val(Text1.Text)
End Sub
Private Sub Form_Initialize()
Call InitCommonControls
End Sub
