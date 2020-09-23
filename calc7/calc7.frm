VERSION 5.00
Begin VB.Form calc7 
   Caption         =   "Super Calculatrice 7.0"
   ClientHeight    =   4290
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9930
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MouseIcon       =   "calc7.frx":0000
   MousePointer    =   4  'Icon
   ScaleHeight     =   4290
   ScaleWidth      =   9930
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command37 
      Caption         =   "."
      Height          =   495
      Left            =   7920
      TabIndex        =   42
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton Command36 
      Caption         =   "Pi"
      Height          =   495
      Left            =   3720
      TabIndex        =   41
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton Command35 
      Caption         =   "x !"
      Height          =   495
      Left            =   3720
      TabIndex        =   40
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton Command34 
      Caption         =   "ln"
      Height          =   495
      Left            =   3120
      TabIndex        =   39
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton Command33 
      Caption         =   "exp"
      Height          =   495
      Left            =   2520
      TabIndex        =   38
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton Command32 
      Caption         =   "rand #"
      Height          =   495
      Left            =   2520
      TabIndex        =   37
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command31 
      Caption         =   "cosh"
      Height          =   495
      Left            =   5160
      TabIndex        =   36
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton Command30 
      Caption         =   "tanh"
      Height          =   495
      Left            =   5760
      TabIndex        =   35
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton Command29 
      Caption         =   "sinh"
      Height          =   495
      Left            =   4560
      TabIndex        =   34
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton Command28 
      Caption         =   "atan"
      Height          =   495
      Left            =   5760
      TabIndex        =   33
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton Command27 
      Caption         =   "acos"
      Height          =   495
      Left            =   5160
      TabIndex        =   32
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton Command26 
      Caption         =   "asin"
      Height          =   495
      Left            =   4560
      TabIndex        =   31
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton Command25 
      Caption         =   "tan"
      Height          =   495
      Left            =   5760
      TabIndex        =   30
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton Command24 
      Caption         =   "cos"
      Height          =   495
      Left            =   5160
      TabIndex        =   29
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton Command23 
      Caption         =   "sin"
      Height          =   495
      Left            =   4560
      TabIndex        =   28
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton Command22 
      Caption         =   "x rt y"
      Height          =   495
      Left            =   3720
      TabIndex        =   27
      ToolTipText     =   "x à la yième racine"
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton Command21 
      Caption         =   "cbrt"
      Height          =   495
      Left            =   3120
      TabIndex        =   26
      ToolTipText     =   "racine cube du nombre affiché à l'écran"
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton Command20 
      Caption         =   "x^y"
      Height          =   495
      Left            =   3720
      TabIndex        =   25
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton Command19 
      Caption         =   "sqrt"
      Height          =   495
      Left            =   2520
      TabIndex        =   24
      ToolTipText     =   "racine carrée du nombre affiché"
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton Command18 
      Caption         =   "x³"
      Height          =   495
      Left            =   3120
      TabIndex        =   23
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton Command17 
      Caption         =   "x²"
      Height          =   495
      Left            =   2520
      TabIndex        =   22
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Vider la liste"
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   3840
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   3570
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton memoire 
      Caption         =   "<------"
      Height          =   375
      Left            =   1800
      TabIndex        =   19
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton Command15 
      Caption         =   "="
      Height          =   495
      Left            =   8520
      TabIndex        =   17
      ToolTipText     =   "égale"
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton Command14 
      Caption         =   "/"
      Height          =   495
      Left            =   9240
      TabIndex        =   16
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton Command13 
      Caption         =   "*"
      Height          =   495
      Left            =   9240
      TabIndex        =   15
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton Command12 
      Caption         =   "-"
      Height          =   495
      Left            =   9240
      TabIndex        =   14
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton Command11 
      Caption         =   "+"
      Height          =   495
      Left            =   9240
      TabIndex        =   13
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton Command10 
      Caption         =   "0"
      Height          =   495
      Left            =   7320
      TabIndex        =   12
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton Command9 
      Caption         =   "9"
      Height          =   495
      Left            =   8520
      TabIndex        =   11
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton Command8 
      Caption         =   "8"
      Height          =   495
      Left            =   7920
      TabIndex        =   10
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      Caption         =   "7"
      Height          =   495
      Left            =   7320
      TabIndex        =   9
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      Caption         =   "6"
      Height          =   495
      Left            =   8520
      TabIndex        =   8
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      Caption         =   "5"
      Height          =   495
      Left            =   7920
      TabIndex        =   7
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "4"
      Height          =   495
      Left            =   7320
      TabIndex        =   6
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "3"
      Height          =   495
      Left            =   8520
      TabIndex        =   5
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "2"
      Height          =   495
      Left            =   7920
      TabIndex        =   4
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   495
      Left            =   7320
      TabIndex        =   3
      Top             =   2640
      Width           =   495
   End
   Begin VB.TextBox affichage 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   0
      Top             =   120
      Width           =   7215
   End
   Begin VB.Label laboper 
      Height          =   255
      Left            =   5040
      TabIndex        =   18
      Top             =   840
      Width           =   615
   End
   Begin VB.Label valeury 
      Caption         =   "Y = 0"
      Height          =   255
      Left            =   5880
      TabIndex        =   2
      Top             =   840
      Width           =   3735
   End
   Begin VB.Label valeurx 
      Caption         =   "X = 0"
      Height          =   255
      Left            =   2640
      TabIndex        =   1
      Top             =   840
      Width           =   2175
   End
   Begin VB.Menu menufichier 
      Caption         =   "Fichier"
      Index           =   1
      NegotiatePosition=   2  'Middle
      Begin VB.Menu mmem 
         Caption         =   "Mettre en mémoire"
         Shortcut        =   ^M
      End
      Begin VB.Menu xmem 
         Caption         =   "Extraire de la mémoire"
         Shortcut        =   ^X
      End
      Begin VB.Menu Programme 
         Caption         =   "Fermer"
         Index           =   2
         Shortcut        =   ^P
      End
   End
   Begin VB.Menu supp 
      Caption         =   "Fonctions supplémentaires"
      Begin VB.Menu graph 
         Caption         =   "Graphique"
         Shortcut        =   ^G
      End
      Begin VB.Menu cbas 
         Caption         =   "Conversions entre bases"
         Shortcut        =   ^B
      End
      Begin VB.Menu ann 
         Caption         =   "Analyse de nombres"
         Shortcut        =   ^A
      End
      Begin VB.Menu qua 
         Caption         =   "Résolution d'équations quadratique"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu aprop 
      Caption         =   "À propos de Super Calculatrice 7.0"
   End
End
Attribute VB_Name = "calc7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private premiernombre As Double
Private secondnombre As Double
Private operation As Variant
Private preteffacer As Boolean

Private Sub Command37_Click()
On Error Resume Next
If preteffacer = True Then affichage.Text = "0"
preteffacer = False
affichage.Text = affichage.Text & "."
End Sub

Private Sub Form_Initialize()
Call InitCommonControls
End Sub
Private Sub ann_Click()
calc7.Hide
Analyse.Show
End Sub

Private Sub aprop_Click()
formprop.Show vbModal
End Sub

Private Sub cbas_Click()
calc7.Hide
Bases.Show
End Sub

Private Sub Command1_Click()
On Error Resume Next
If preteffacer = True Then affichage.Text = ""
preteffacer = False
affichage.Text = affichage.Text & "1"
End Sub

Private Sub Command10_Click()
On Error Resume Next
If preteffacer = True Then affichage.Text = ""
preteffacer = False
affichage.Text = affichage.Text & "0"
End Sub

Private Sub Command11_Click()
On Error Resume Next
valeurx.Caption = affichage.Text
operation = "addition"
premiernombre = affichage.Text
preteffacer = True
laboper.Caption = "+"
End Sub

Private Sub Command12_Click()
On Error Resume Next
valeurx.Caption = affichage.Text
operation = "soustraction"
premiernombre = affichage.Text
preteffacer = True
laboper.Caption = "-"
End Sub

Private Sub Command13_Click()
On Error Resume Next
valeurx.Caption = affichage.Text
operation = "multiplication"
premiernombre = affichage.Text
preteffacer = True
laboper.Caption = "*"
End Sub

Private Sub Command14_Click()
On Error Resume Next
valeurx.Caption = affichage.Text
operation = "division"
premiernombre = affichage.Text
preteffacer = True
laboper.Caption = "/"
End Sub

Private Sub Command15_Click()
On Error GoTo erreurdiv
valeury.Caption = affichage.Text
secondnombre = affichage.Text
If operation = "addition" Then affichage.Text = premiernombre + secondnombre
If operation = "soustraction" Then affichage.Text = premiernombre - secondnombre
If operation = "multiplication" Then affichage.Text = premiernombre * secondnombre
If operation = "division" Then affichage.Text = premiernombre / secondnombre
If operation = "exposant" Then affichage.Text = premiernombre ^ secondnombre
If operation = "racine" Then affichage.Text = premiernombre ^ (1 / secondnombre)

GoTo effacer
erreurdiv:
Dim x As VbMsgBoxResult
x = MsgBox(Err.Description, vbOKOnly, "Super Calculatrice 7.0")
effacer:
preteffacer = True
End Sub

Private Sub Command16_Click()
On Error Resume Next
List1.Clear
End Sub

Private Sub Command17_Click()
On Error Resume Next
preteffacer = True
affichage.Text = Val(affichage.Text) ^ 2
End Sub

Private Sub Command18_Click()
On Error Resume Next
preteffacer = True
affichage.Text = Val(affichage.Text) ^ 3
End Sub

Private Sub Command19_Click()
On Error Resume Next
preteffacer = True
affichage.Text = Val(affichage.Text) ^ (1 / 2)
End Sub

Private Sub Command2_Click()
On Error Resume Next
If preteffacer = True Then affichage.Text = ""
preteffacer = False
affichage.Text = affichage.Text & "2"
End Sub

Private Sub Command20_Click()
On Error Resume Next
valeurx.Caption = affichage.Text
operation = "exposant"
premiernombre = affichage.Text
preteffacer = True
laboper.Caption = "^"
End Sub

Private Sub Command21_Click()
On Error Resume Next
preteffacer = True
affichage.Text = Val(affichage.Text) ^ (1 / 3)
End Sub

Private Sub Command22_Click()
On Error Resume Next
valeurx.Caption = affichage.Text
operation = "racine"
premiernombre = affichage.Text
preteffacer = True
laboper.Caption = "^ 1 /"
End Sub

Private Sub Command23_Click()
On Error Resume Next
preteffacer = True
affichage.Text = Sin(Val(affichage.Text))
End Sub

Private Sub Command24_Click()
On Error Resume Next
preteffacer = True
affichage.Text = Cos(Val(affichage.Text))
End Sub

Private Sub Command25_Click()
On Error Resume Next
preteffacer = True
affichage.Text = Tan(Val(affichage.Text))
End Sub

Private Sub Command26_Click()
On Error Resume Next
preteffacer = True
affichage.Text = (Atn(Val(affichage.Text) / Sqr(-(Val(affichage.Text)) * Val(affichage.Text) + 1)))
End Sub

Private Sub Command27_Click()
On Error Resume Next
preteffacer = True
affichage.Text = Atn(-(Val(affichage.Text)) / Sqr(-(Val(affichage.Text)) * Val(affichage.Text) + 1)) + 2 * Atn(1)
End Sub

Private Sub Command28_Click()
On Error Resume Next
preteffacer = True
affichage.Text = Atn(Val(affichage.Text))
End Sub

Private Sub Command29_Click()
On Error Resume Next
preteffacer = True
Dim x As String
x = affichage.Text
affichage.Text = (Exp(x) - Exp(-x)) / 2
End Sub

Private Sub Command3_Click()
On Error Resume Next
If preteffacer = True Then affichage.Text = ""
preteffacer = False
affichage.Text = affichage.Text & "3"
End Sub

Private Sub Command30_Click()
On Error Resume Next
preteffacer = True
Dim x As String
x = affichage.Text
affichage.Text = (Exp(x) - Exp(-x)) / (Exp(x) + Exp(-x))
End Sub

Private Sub Command31_Click()
On Error Resume Next
preteffacer = True
Dim x As String
x = affichage.Text
affichage.Text = (Exp(x) + Exp(-x)) / 2
End Sub

Private Sub Command32_Click()
On Error Resume Next
preteffacer = True
affichage.Text = Rnd
End Sub

Private Sub Command33_Click()
On Error Resume Next
preteffacer = True
affichage.Text = Exp(Val(affichage.Text))
End Sub

Private Sub Command34_Click()
On Error Resume Next
preteffacer = True
affichage.Text = Log(Val(affichage.Text))
End Sub

Private Sub Command35_Click()
On Error Resume Next
preteffacer = True
affichage.Text = factorielle(Val(affichage.Text))
End Sub

Private Sub Command36_Click()
On Error Resume Next
preteffacer = True
affichage.Text = "3.1415926535"
End Sub

Private Sub Command4_Click()
On Error Resume Next
If preteffacer = True Then affichage.Text = ""
preteffacer = False
affichage.Text = affichage.Text & "4"
End Sub

Private Sub Command5_Click()
On Error Resume Next
If preteffacer = True Then affichage.Text = ""
preteffacer = False
affichage.Text = affichage.Text & "5"
End Sub

Private Sub Command6_Click()
On Error Resume Next
If preteffacer = True Then affichage.Text = ""
preteffacer = False
affichage.Text = affichage.Text & "6"
End Sub

Private Sub Command7_Click()
On Error Resume Next
If preteffacer = True Then affichage.Text = ""
preteffacer = False
affichage.Text = affichage.Text & "7"
End Sub

Private Sub Command8_Click()
On Error Resume Next
If preteffacer = True Then affichage.Text = ""
preteffacer = False
affichage.Text = affichage.Text & "8"
End Sub

Private Sub Command9_Click()
On Error Resume Next
If preteffacer = True Then affichage.Text = ""
preteffacer = False
affichage.Text = affichage.Text & "9"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Programme_Click (0)
End Sub


Private Sub graph_Click()
calc7.Hide
GraphPlot.Show
End Sub

Private Sub List1_DblClick()
On Error Resume Next
affichage.Text = List1.List(List1.ListIndex)
preteffacer = True
End Sub

Private Sub memoire_Click()
On Error Resume Next
If affichage.Text = "" Then GoTo 10
List1.AddItem (affichage.Text)
10 End Sub

Private Sub mmem_Click()
On Error Resume Next
If affichage.Text = "" Then GoTo 10
List1.AddItem (affichage.Text)
10 End Sub

Private Sub Programme_Click(Index As Integer)
On Error Resume Next
End
End Sub

Private Sub qua_Click()
calc7.Hide
frmQuadratic.Show
End Sub

Private Sub xmem_Click()
On Error Resume Next
affichage.Text = List1.List(List1.ListIndex)
preteffacer = True
End Sub
