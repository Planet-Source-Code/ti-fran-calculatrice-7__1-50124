VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About "
   ClientHeight    =   3135
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5655
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2163.833
   ScaleMode       =   0  'User
   ScaleWidth      =   5310.337
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Quit"
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H80000009&
      Caption         =   $"frmAbout.frx":0442
      ForeColor       =   &H00000000&
      Height          =   690
      Left            =   960
      TabIndex        =   0
      Top             =   1080
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H80000009&
      Caption         =   "Simple Graphic Plotter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1050
      TabIndex        =   1
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1697.936
      Y2              =   1697.936
   End
   Begin VB.Label lblVersion 
      BackColor       =   &H80000009&
      Caption         =   "Version 1.0"
      Height          =   225
      Left            =   1080
      TabIndex        =   2
      Top             =   720
      Width           =   3405
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Command1_Click()
Unload Me
End Sub
Private Sub Form_Initialize()
Call InitCommonControls
End Sub
