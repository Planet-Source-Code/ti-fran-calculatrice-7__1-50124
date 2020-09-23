VERSION 5.00
Begin VB.Form divnombre 
   Caption         =   "Super Calculatrice 7.0"
   ClientHeight    =   3690
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   4665
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3690
   ScaleWidth      =   4665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Quitter"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3240
      Width           =   4455
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   4455
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Trouver les diviseurs"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Entre un nombre"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "divnombre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo erreur
If Val(Text1.Text) > 10000000 Then GoTo 33
List1.Clear
Dim a As Long
Dim i As Long
a = Val(Text1)
i = 0
Dim u As Long
u = 0
Do While i < a
i = i + 1
If a Mod i = 0 Then GoTo 5 Else GoTo 50
5 List1.AddItem (i)
u = u + 1
50 Loop
Text2.Text = "Ce nombre a " & u & " diviseurs"
If u = 2 Then Text2 = "Ce nombre est premier "
Exit Sub
erreur:
MsgBox Err.Description
33 End Sub

Private Sub Command2_Click()
divnombre.Hide
calc7.Show
End Sub

Private Sub quitter_Click()
divnombre.Hide
calc7.Show
End Sub

Private Sub Form_Initialize()
Call InitCommonControls
End Sub
