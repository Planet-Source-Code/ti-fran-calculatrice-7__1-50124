VERSION 5.00
Begin VB.Form factprem 
   Caption         =   "Super Calculatrice 7.0"
   ClientHeight    =   3555
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3555
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Quitter"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3120
      Width           =   4335
   End
   Begin VB.ListBox LstRes 
      Height          =   1815
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Décomposer en facteurs premiers"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   4335
   End
   Begin VB.TextBox txtnombre 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "Entre un nombre"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "factprem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo erreur
If Val(txtnombre.Text) > 100000 Then GoTo 33

    ' Fait la décomposition
    Dim Facteurs() As Long
    Facteurs = Decompose(txtnombre)
    
    ' Remplit la liste
    Dim n
    LstRes.Clear
    For Each n In Facteurs
        LstRes.AddItem n
    Next
    Exit Sub
erreur:
Dim xr As VbMsgBoxResult
    xr = MsgBox(Err.Description, vbExclamation, erreur)
33 End Sub

Private Sub Form_Initialize()
Call InitCommonControls
End Sub
Private Sub Command2_Click()
factprem.Hide
calc7.Show
End Sub

