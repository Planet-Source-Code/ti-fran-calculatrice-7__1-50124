VERSION 5.00
Begin VB.Form frmQuadratic 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Super Calculatrice 7.0"
   ClientHeight    =   6930
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   9345
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   9345
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmndExit 
      Caption         =   "Sor&tir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   6
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmndClear 
      Caption         =   "Eff&acer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   5
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmndCalculate 
      Caption         =   "&Calculer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   5880
      Width           =   1215
   End
   Begin VB.PictureBox picGraph 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   4680
      ScaleHeight     =   4275
      ScaleMode       =   0  'User
      ScaleWidth      =   4395
      TabIndex        =   15
      Top             =   2400
      Width           =   4455
   End
   Begin VB.ListBox lstTable 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   4680
      TabIndex        =   7
      Top             =   480
      Width           =   4455
   End
   Begin VB.TextBox txtC 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3000
      TabIndex        =   3
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox txtB 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1680
      TabIndex        =   2
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox txtA 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Entrez l'équation ci-dessous."
      Height          =   375
      Left            =   240
      TabIndex        =   24
      Top             =   720
      Width           =   4095
   End
   Begin VB.Label lblEquation 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "2A"
      Height          =   300
      Index           =   1
      Left            =   1680
      TabIndex        =   23
      Top             =   5400
      Width           =   1860
   End
   Begin VB.Line lneEquation 
      X1              =   1560
      X2              =   3600
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Label lblFX 
      Caption         =   "F(x) ="
      Height          =   300
      Left            =   840
      TabIndex        =   22
      Top             =   5160
      Width           =   600
   End
   Begin VB.Label lblEquation 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "-B ± sqr(B² - 4AC)"
      Height          =   300
      Index           =   0
      Left            =   1635
      TabIndex        =   21
      Top             =   4920
      Width           =   1890
   End
   Begin VB.Label lblGraph 
      AutoSize        =   -1  'True
      Caption         =   "Graphique"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6720
      TabIndex        =   20
      Top             =   2040
      Width           =   945
   End
   Begin VB.Label lblDiscriminant 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   2400
      TabIndex        =   19
      Top             =   3840
      Width           =   1920
   End
   Begin VB.Label lblSolutions 
      AutoSize        =   -1  'True
      Caption         =   "Discriminant:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   2400
      TabIndex        =   18
      Top             =   3480
      Width           =   1140
   End
   Begin VB.Label lblXVertex 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   240
      TabIndex        =   17
      Top             =   3840
      Width           =   1920
   End
   Begin VB.Label lblSolutions 
      AutoSize        =   -1  'True
      Caption         =   "X Vertex:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   240
      TabIndex        =   16
      Top             =   3480
      Width           =   780
   End
   Begin VB.Label lblXY 
      Caption         =   "(x, y)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6960
      TabIndex        =   14
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblYIntercept 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   2400
      TabIndex        =   13
      Top             =   2640
      Width           =   1920
   End
   Begin VB.Label lblXintercepts 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   240
      TabIndex        =   12
      Top             =   2640
      Width           =   1920
   End
   Begin VB.Label lblSolutions 
      AutoSize        =   -1  'True
      Caption         =   "Y Intercept:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   2400
      TabIndex        =   11
      Top             =   2400
      Width           =   990
   End
   Begin VB.Label lblSolutions 
      AutoSize        =   -1  'True
      Caption         =   "X Intercepts:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Top             =   2400
      Width           =   1080
   End
   Begin VB.Label lblC 
      AutoSize        =   -1  'True
      Caption         =   "= 0"
      Height          =   300
      Left            =   3960
      TabIndex        =   9
      Top             =   1200
      Width           =   330
   End
   Begin VB.Label lblB 
      AutoSize        =   -1  'True
      Caption         =   "X +"
      Height          =   300
      Left            =   2640
      TabIndex        =   8
      Top             =   1200
      Width           =   360
   End
   Begin VB.Label lblA 
      AutoSize        =   -1  'True
      Caption         =   "X² +"
      Height          =   300
      Left            =   1200
      TabIndex        =   0
      Top             =   1200
      Width           =   435
   End
   Begin VB.Menu quit 
      Caption         =   "Quitter"
      Index           =   1
   End
End
Attribute VB_Name = "frmQuadratic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmndCalculate_Click()

    ' Variables for Coefficients
    Dim sngA As Single
    Dim sngB As Single
    Dim sngC As Single
    
    ' Variable to Hold Discriminant
    Dim Disc As Single
    
    ' Variables for Intercepts and Vertex
    Dim sngX1 As Single, sngX2 As Single
    Dim sngXVertex As Single
        
    ' Get Coefficient Values
    sngA = Val(txtA.Text)
    sngB = Val(txtB.Text)
    sngC = Val(txtC.Text)
    
    ' Make sure User entered Quadratic Equation
    If sngA = 0 Then
        MsgBox "Not a quadratic equation!", vbExclamation, "Error"
        txtA.SetFocus
        Exit Sub
    End If
    
    ' Calculate Y Intercept
    lblYIntercept.Caption = sngC
        
    ' Calculate X Vertex
    Dim sngH As Single
    sngH = (-sngB) / (2 * sngA)
    sngXVertex = (sngA * sngH * sngH) + (sngB * sngH) + sngC
    lblXVertex.Caption = sngXVertex
   
    ' Calculate Discriminant
    Disc = (sngB * sngB) - (4 * sngA * sngC)
    lblDiscriminant.Caption = Disc
    ' Show Equation
    lblEquation(0).Caption = -sngB & " ± sqr(" & Disc & ")"
    lblEquation(1).Caption = 2 * sngA
         
    ' Calculate X Intercept(s)
    If Disc < 0 Then
        ' Variables for Complex Numbers
        Dim sngReal As Single
        Dim sngImaginary As Single
            
        sngReal = (-sngB) / (2 * sngA)
        sngImaginary = (Sqr(Abs(Disc))) / (2 * sngA)
                           
        lblXintercepts.Caption = sngReal & " ± " & sngImaginary & "i"
    ElseIf Disc = 0 Then
        sngX1 = (-sngB + Sqr(Disc)) / (2 * sngA)
        
        lblXintercepts.Caption = "Tangent: " & sngX1
    ElseIf Disc > 0 Then
        sngX1 = (-sngB + Sqr(Disc)) / (2 * sngA)
        sngX2 = (-sngB - Sqr(Disc)) / (2 * sngA)
        
        lblXintercepts.Caption = sngX1 & ", " & sngX2
    End If
    
    ' Generate Table Values
    Dim x As Single, y As Single
    Dim PointNumber As Integer
    Const INCREMENT = 0.01
    ReDim YVals(1 To 21 / INCREMENT) As Single
    PointNumber = 1
    For x = -10 To 10 Step INCREMENT
        y = (sngA * x * x) + (sngB * x) + sngC
        YVals(PointNumber) = y
        PointNumber = PointNumber + 1
        lstTable.AddItem Format(x, "Fixed") & vbTab & Format(y, "Fixed")
    Next x
    
    ' Set Up Graph
    picGraph.Scale (-10, 20)-(10, -20)
    ' Draw x and y axis
    picGraph.Line (0, -20)-(0, 20), RGB(0, 200, 0)
    picGraph.Line (-10, 0)-(10, 0), RGB(0, 200, 0)
    
    PointNumber = 1
    ' Plot Graph
    For x = -10 To 10 Step INCREMENT
        picGraph.PSet (x, YVals(PointNumber))
        PointNumber = PointNumber + 1
    Next x
    
End Sub
Private Sub cmndClear_Click()

    lstTable.Clear
    
    txtA.Text = ""
    txtB.Text = ""
    txtC.Text = ""
    
    lblXintercepts.Caption = ""
    lblYIntercept.Caption = ""
    lblXVertex.Caption = ""
    lblDiscriminant.Caption = ""
    lblEquation(0).Caption = "-B ± sqr(B² - 4AC)"
    lblEquation(1).Caption = "2A"
    
    picGraph.Cls
    
    txtA.SetFocus
    
End Sub
Private Sub cmndExit_Click()
frmQuadratic.Hide
calc7.Show
End Sub
Private Sub Form_Initialize()
Call InitCommonControls
End Sub
Private Sub quit_Click(Index As Integer)
frmQuadratic.Hide
calc7.Show
End Sub
