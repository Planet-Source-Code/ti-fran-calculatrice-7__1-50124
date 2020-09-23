VERSION 5.00
Begin VB.Form Trigo 
   Caption         =   "Mode trigonométrique"
   ClientHeight    =   2190
   ClientLeft      =   2775
   ClientTop       =   1245
   ClientWidth     =   6945
   ControlBox      =   0   'False
   Icon            =   "Trigo.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   2190
   ScaleWidth      =   6945
   Begin VB.CommandButton Command3 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   3000
      TabIndex        =   25
      ToolTipText     =   "Close this window"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Effacer tout"
      Height          =   375
      Left            =   4440
      TabIndex        =   24
      ToolTipText     =   "Reset the plotter to the default value"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Dessiner"
      Height          =   375
      Left            =   1560
      TabIndex        =   23
      ToolTipText     =   "Press this key to draw"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   6300
      TabIndex        =   17
      Text            =   "1"
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   4845
      TabIndex        =   16
      Text            =   "0"
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   3360
      TabIndex        =   15
      Text            =   "0"
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   2100
      TabIndex        =   14
      Text            =   "0"
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   720
      TabIndex        =   13
      Text            =   "0"
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   6300
      TabIndex        =   12
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   4845
      TabIndex        =   11
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   3360
      TabIndex        =   10
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2100
      TabIndex        =   9
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   720
      TabIndex        =   8
      ToolTipText     =   "Enter the value of coefficient here"
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label13 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4275
      TabIndex        =   22
      Top             =   960
      Width           =   135
   End
   Begin VB.Label Label12 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4275
      TabIndex        =   21
      Top             =   360
      Width           =   135
   End
   Begin VB.Label Label11 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1590
      TabIndex        =   20
      Top             =   960
      Width           =   135
   End
   Begin VB.Label Label10 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1590
      TabIndex        =   19
      Top             =   360
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "sin  x+"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   600
      X2              =   6720
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "y ="
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   18
      Top             =   690
      Width           =   735
   End
   Begin VB.Label Label9 
      Caption         =   "cos x+"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5355
      TabIndex        =   7
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "cos  x+"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   6
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "sin x+"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "sin  x+"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "cos x+"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5355
      TabIndex        =   3
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "cos  x+"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "sin x+"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "Trigo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Dim C1 As Double, C2 As Double
Dim equ1 As String, equ2 As String
Dim i As Integer
Dim d(9) As Integer
Call RndColor

With wait
    .ProgressBar1.Min = LBound(d)
    .ProgressBar1.Max = UBound(d)
    .ProgressBar1.Visible = True
    .ProgressBar1.Value = ProgressBar1.Min
End With

With GraphPlot
    .Picture1.Scale (-10, 10)-(10, -10)   'Scale
    .Picture1.ForeColor = RGB(M, n, O)
    .Text4 = Val(GraphPlot.Text4) - 5
    .Picture2.Line (10, Val(GraphPlot.Text4.Text))-(-10, Val(GraphPlot.Text4.Text)), RGB(M, n, O)
End With

For i = 0 To 9
    d(i) = Val(Text1(i).Text)
Next i
  
wait.Show
GraphPlot.StatusBar1.Panels(3).Text = "Please wait while drawing......"
For C1 = -10 To 10 Step 0.001
    C2 = ((d(0) * Sin(C1)) ^ 2 + d(1) * Sin(C1) + d(2) * (Cos(C1)) ^ 2 + d(3) * Cos(C1) + d(4)) / (d(5) * (Sin(C1)) ^ 2 + d(6) * Sin(C1) + (d(7) * Cos(C1)) ^ 2 + d(8) * Cos(C1) + d(9))
    GraphPlot.Picture1.PSet (C1, C2) 'draws the graph
    wait.ProgressBar1.Value = C1
Next C1
wait.Hide
GraphPlot.StatusBar1.Panels(3).Text = ""
wait.ProgressBar1.Visible = False
wait.ProgressBar1.Value = ProgressBar1.Min

equ1 = "    " & d(0) & "(SinX)^2+" & d(1) & "SinX+" & d(2) & "(CosX)^2+" & d(3) & "CosX+" & _
        d(4) & "    " & d(5) & "(SinX)^2+" & d(6) & "SinX+(" & d(7) & "CosX)^2+" & d(8) & "CosX+" & d(9)


equ2 = "y=" & d(0) & "(SinX)^2+" & d(1) & "SinX+" & d(2) & "(CosX)^2+" & d(3) & "CosX+" & d(4)

If Val(GraphPlot.Text5) <= 12 Then
    
   GraphPlot.Label1(GraphPlot.Text5.Text).ForeColor = RGB(M, n, O)
   If d(0) = 0 And d(1) = 0 And d(2) = 0 And d(3) = 0 And d(4) = 0 Then
        GraphPlot.Label1(GraphPlot.Text5.Text).Caption = "0"
    ElseIf d(5) = 0 And d(6) = 0 And d(7) = 0 And d(8) = 0 And d(9) = 1 Then
        GraphPlot.Label1(GraphPlot.Text5.Text).Caption = equ2
    Else
        GraphPlot.Line2(GraphPlot.Text5.Text).Visible = True
        GraphPlot.Label1(GraphPlot.Text5.Text).Caption = equ1
    End If
        GraphPlot.Text5 = Val(GraphPlot.Text5) + 1
Else
Call sort
    GraphPlot.Label1(12).ForeColor = RGB(M, n, O)
    If d(0) = 0 And d(1) = 0 And d(2) = 0 And d(3) = 0 And d(4) = 0 Then
        GraphPlot.Label1(GraphPlot.Text5.Text).Caption = "y = 0"
    ElseIf d(5) = 0 And d(6) = 0 And d(7) = 0 And d(8) = 0 And d(9) = 1 Then
        GraphPlot.Label1(GraphPlot.Text5.Text).Caption = equ2
    Else
        GraphPlot.Line2(GraphPlot.Text5.Text).Visible = True
        GraphPlot.Label1(GraphPlot.Text5.Text).Caption = equ1
    End If
End If
End Sub
Private Sub Command2_Click()
With GraphPlot
    .Text1(8).Text = 10
    .Picture1.Scale (-10, 10)-(10, -10)
    .Picture1.Cls 'clear picture box
    .Picture2.Cls

    .Text4.Text = 103 'reset the index
    .Text5.Text = 0

For i = 0 To 12
    .Label1(i).Caption = ""
Next i

For i = 0 To 6
    .Text1(i).Text = ""
Next i



.Text1(7) = 1

.Text2.Visible = False
.Text3.Visible = False
.Label12.Visible = False
.Label19.Visible = False
End With

GraphPlot.Picture1.Line (0, -10)-(0, 10), QBColor(0) 'Draw y-axis
GraphPlot.Picture1.Line (-10, 0)-(10, 0), QBColor(0) 'Draw x-axis

For i1 = -10 To 10
    GraphPlot.Picture1.Line (-0.1, i1)-(0.1, i1), QBColor(0) 'Draw index
    GraphPlot.Picture1.Line (i1, -0.1)-(i1, 0.1), QBColor(0)
Next i1

For i = 0 To 12
    GraphPlot.Line2(i).Visible = False
Next i

For i = 0 To 8
    Text1(i).Text = ""
Next i

Text1(9) = 1

End Sub

Private Sub Command3_Click()
Ftrigo = False
GraphPlot.Label5.ToolTipText = "Mouse x-coordinate"
Unload Me
End Sub

Private Sub Form_Initialize()
Call InitCommonControls
End Sub
Private Sub Form_Load()
    Randomize
    With GraphPlot
    .Text2.Visible = False
    .Text3.Visible = False
    .Label12.Visible = False
    .Label19.Visible = False
    .Label5.ToolTipText = "Mouse coordinate in degree (blue)"
    End With
    
    For i = 0 To 3
        Text1(i).ToolTipText = "Enter the value of coefficient here"
    Next i
    Text1(4).ToolTipText = "Enter a contant here"
    Text1(9).ToolTipText = "Enter a contant here"
    For i = 5 To 8
        Text1(i).ToolTipText = "Enter the value of coefficient here"
    Next i
End Sub
