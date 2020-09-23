VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form GraphPlot 
   AutoRedraw      =   -1  'True
   Caption         =   "Super Calculatrice 7.0"
   ClientHeight    =   7995
   ClientLeft      =   150
   ClientTop       =   735
   ClientWidth     =   11880
   ControlBox      =   0   'False
   Icon            =   "Graph.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7995
   ScaleWidth      =   11880
   Begin MSComctlLib.StatusBar StatusBar1 
      Height          =   355
      Left            =   75
      TabIndex        =   55
      Top             =   7620
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "17:02"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            AutoSize        =   2
            Object.Width           =   2302
            MinWidth        =   2293
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5644
            MinWidth        =   5644
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   0
            Enabled         =   0   'False
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   0
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   2
            Text            =   "Thank you for using graphic plotter!"
            TextSave        =   "2002-12-16"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00E0E0E0&
      Height          =   6135
      Left            =   7970
      ScaleHeight     =   6075
      ScaleWidth      =   795
      TabIndex        =   31
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   380
      Index           =   8
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   24
      Text            =   "10"
      ToolTipText     =   "Value of current scale"
      Top             =   10
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   7
      Left            =   3840
      TabIndex        =   23
      Text            =   "1"
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   6
      Left            =   2880
      TabIndex        =   22
      Text            =   "0"
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   5
      Left            =   1920
      TabIndex        =   21
      Text            =   "0"
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2160
      TabIndex        =   19
      Top             =   1000
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1200
      TabIndex        =   18
      Top             =   1000
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   4
      Left            =   960
      TabIndex        =   17
      Text            =   "0"
      Top             =   600
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Reset"
      Height          =   375
      Left            =   6480
      TabIndex        =   16
      ToolTipText     =   "Reset the plotter to the default value"
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   3
      Left            =   3840
      TabIndex        =   11
      Top             =   165
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   2
      Left            =   2880
      TabIndex        =   10
      Top             =   165
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   1920
      TabIndex        =   9
      Top             =   165
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   960
      TabIndex        =   8
      Top             =   165
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Draw"
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      ToolTipText     =   "Press this key to draw"
      Top             =   600
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   6135
      Left            =   0
      ScaleHeight     =   6075
      ScaleWidth      =   7875
      TabIndex        =   0
      Top             =   1320
      Width           =   7935
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   480
         TabIndex        =   51
         Text            =   "0"
         Top             =   3240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   480
         TabIndex        =   50
         Text            =   "103"
         Top             =   2760
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin VB.Label Label26 
      Caption         =   ","
      Height          =   255
      Left            =   6840
      TabIndex        =   54
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Label25 
      Caption         =   ")"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7365
      TabIndex        =   53
      Top             =   1080
      Width           =   135
   End
   Begin VB.Label Label24 
      Caption         =   "("
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   52
      Top             =   1080
      Width           =   135
   End
   Begin VB.Line Line1 
      X1              =   840
      X2              =   4440
      Y1              =   525
      Y2              =   525
   End
   Begin VB.Line Line2 
      Index           =   12
      X1              =   9000
      X2              =   11865
      Y1              =   7275
      Y2              =   7275
   End
   Begin VB.Line Line2 
      Index           =   11
      X1              =   9000
      X2              =   11865
      Y1              =   6795
      Y2              =   6795
   End
   Begin VB.Line Line2 
      Index           =   10
      X1              =   9000
      X2              =   11865
      Y1              =   6315
      Y2              =   6315
   End
   Begin VB.Line Line2 
      Index           =   9
      X1              =   9000
      X2              =   11865
      Y1              =   5835
      Y2              =   5835
   End
   Begin VB.Line Line2 
      Index           =   8
      X1              =   9000
      X2              =   11865
      Y1              =   5355
      Y2              =   5355
   End
   Begin VB.Line Line2 
      Index           =   7
      X1              =   9000
      X2              =   11865
      Y1              =   4875
      Y2              =   4875
   End
   Begin VB.Line Line2 
      Index           =   6
      X1              =   9000
      X2              =   11865
      Y1              =   4395
      Y2              =   4395
   End
   Begin VB.Line Line2 
      Index           =   5
      X1              =   9000
      X2              =   11865
      Y1              =   3915
      Y2              =   3915
   End
   Begin VB.Line Line2 
      Index           =   4
      X1              =   9000
      X2              =   11865
      Y1              =   3435
      Y2              =   3435
   End
   Begin VB.Line Line2 
      Index           =   3
      X1              =   9000
      X2              =   11865
      Y1              =   2955
      Y2              =   2955
   End
   Begin VB.Line Line2 
      Index           =   2
      X1              =   9000
      X2              =   11865
      Y1              =   2475
      Y2              =   2475
   End
   Begin VB.Line Line2 
      Index           =   1
      X1              =   9000
      X2              =   11865
      Y1              =   1995
      Y2              =   1995
   End
   Begin VB.Line Line2 
      Index           =   0
      X1              =   9000
      X2              =   11865
      Y1              =   1515
      Y2              =   1515
   End
   Begin VB.Shape Shape13 
      Height          =   390
      Left            =   8865
      Top             =   7080
      Width           =   3045
   End
   Begin VB.Shape Shape12 
      Height          =   390
      Left            =   8865
      Top             =   6600
      Width           =   3045
   End
   Begin VB.Shape Shape11 
      Height          =   390
      Left            =   8865
      Top             =   6120
      Width           =   3045
   End
   Begin VB.Shape Shape10 
      Height          =   390
      Left            =   8865
      Top             =   5640
      Width           =   3045
   End
   Begin VB.Shape Shape9 
      Height          =   390
      Left            =   8865
      Top             =   5160
      Width           =   3045
   End
   Begin VB.Shape Shape8 
      Height          =   390
      Left            =   8865
      Top             =   4680
      Width           =   3045
   End
   Begin VB.Shape Shape7 
      Height          =   390
      Left            =   8865
      Top             =   4200
      Width           =   3045
   End
   Begin VB.Shape Shape6 
      Height          =   390
      Left            =   8865
      Top             =   3720
      Width           =   3045
   End
   Begin VB.Shape Shape5 
      Height          =   390
      Left            =   8865
      Top             =   3240
      Width           =   3045
   End
   Begin VB.Shape Shape4 
      Height          =   390
      Left            =   8865
      Top             =   2760
      Width           =   3045
   End
   Begin VB.Shape Shape3 
      Height          =   390
      Left            =   8865
      Top             =   2280
      Width           =   3045
   End
   Begin VB.Shape Shape2 
      Height          =   390
      Left            =   8865
      Top             =   1800
      Width           =   3045
   End
   Begin VB.Shape Shape1 
      Height          =   390
      Left            =   8865
      Top             =   1320
      Width           =   3045
   End
   Begin VB.Label Label23 
      Caption         =   "equations history"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   9360
      TabIndex        =   49
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      Caption         =   "Last 13"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   9720
      TabIndex        =   48
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   390
      Index           =   12
      Left            =   8865
      TabIndex        =   47
      Top             =   7080
      Width           =   3045
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   390
      Index           =   11
      Left            =   8865
      TabIndex        =   46
      Top             =   6600
      Width           =   3045
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   390
      Index           =   10
      Left            =   8865
      TabIndex        =   45
      Top             =   6120
      Width           =   3045
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   390
      Index           =   0
      Left            =   8865
      TabIndex        =   44
      Top             =   1320
      Width           =   3045
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   390
      Index           =   1
      Left            =   8865
      TabIndex        =   43
      Top             =   1800
      Width           =   3045
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   390
      Index           =   2
      Left            =   8865
      TabIndex        =   42
      Top             =   2280
      Width           =   3045
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   390
      Index           =   3
      Left            =   8865
      TabIndex        =   41
      Top             =   2760
      Width           =   3045
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   390
      Index           =   4
      Left            =   8865
      TabIndex        =   40
      Top             =   3240
      Width           =   3045
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   390
      Index           =   5
      Left            =   8865
      TabIndex        =   39
      Top             =   3720
      Width           =   3045
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   390
      Index           =   6
      Left            =   8865
      TabIndex        =   38
      Top             =   4200
      Width           =   3045
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   390
      Index           =   7
      Left            =   8865
      TabIndex        =   37
      Top             =   4680
      Width           =   3045
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   390
      Index           =   8
      Left            =   8865
      TabIndex        =   36
      Top             =   5160
      Width           =   3045
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   390
      Index           =   9
      Left            =   8865
      TabIndex        =   35
      Top             =   5640
      Width           =   3045
   End
   Begin VB.Label Label22 
      Caption         =   "Current Scale:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   34
      Top             =   10
      Width           =   1335
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      Caption         =   "Index"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   33
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      Caption         =   "Line"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   32
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label19 
      Caption         =   "This equation has no real roots"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   30
      Top             =   1005
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label Label17 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   29
      Top             =   600
      Width           =   135
   End
   Begin VB.Label Label16 
      Caption         =   "x +"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   28
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label15 
      Caption         =   "x +"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   27
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label14 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   26
      Top             =   600
      Width           =   135
   End
   Begin VB.Label Label13 
      Caption         =   "x +"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   25
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label12 
      Caption         =   "Roots:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   20
      Top             =   1005
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   13
      Top             =   165
      Width           =   135
   End
   Begin VB.Label Label6 
      Caption         =   "x +"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   165
      Width           =   615
   End
   Begin VB.Label Label11 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   15
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "y"
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
      Left            =   120
      TabIndex        =   14
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label8 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   12
      Top             =   165
      Width           =   135
   End
   Begin VB.Label Label7 
      Caption         =   "x +"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   7
      Top             =   165
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "x +"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   5
      Top             =   165
      Width           =   495
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   6960
      TabIndex        =   4
      ToolTipText     =   "Mouse y-coordinate"
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label5 
      Height          =   255
      Left            =   6480
      TabIndex        =   3
      ToolTipText     =   "Mouse x-coordinate"
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Coordinate:"
      Height          =   255
      Left            =   5400
      TabIndex        =   2
      ToolTipText     =   "Mouse coordinate"
      Top             =   1080
      Width           =   855
   End
   Begin VB.Menu actionItem 
      Caption         =   "&Action"
      Begin VB.Menu triItem 
         Caption         =   "&Trigonometic Mode"
      End
      Begin VB.Menu clearItem 
         Caption         =   "&Reset"
      End
      Begin VB.Menu sp1 
         Caption         =   "-"
      End
      Begin VB.Menu aboutme 
         Caption         =   "Ab&out"
      End
      Begin VB.Menu sp2 
         Caption         =   "-"
      End
      Begin VB.Menu exitItem 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu ScaleItem 
      Caption         =   "&Scale"
      Begin VB.Menu item10 
         Caption         =   "&10"
      End
      Begin VB.Menu item20 
         Caption         =   "&20"
      End
      Begin VB.Menu item30 
         Caption         =   "&30"
      End
      Begin VB.Menu item40 
         Caption         =   "&40"
      End
      Begin VB.Menu item50 
         Caption         =   "&50"
      End
      Begin VB.Menu item60 
         Caption         =   "&60"
      End
      Begin VB.Menu item70 
         Caption         =   "&70"
      End
      Begin VB.Menu item80 
         Caption         =   "&80"
      End
      Begin VB.Menu item90 
         Caption         =   "&90"
      End
      Begin VB.Menu item100 
         Caption         =   "1&00"
      End
   End
   Begin VB.Menu copyItem 
      Caption         =   "C&opy Equation"
      Begin VB.Menu eq1 
         Caption         =   "Copy &1"
      End
      Begin VB.Menu eq2 
         Caption         =   "Copy &2"
      End
      Begin VB.Menu eq3 
         Caption         =   "Copy &3"
      End
      Begin VB.Menu eq4 
         Caption         =   "Copy &4"
      End
      Begin VB.Menu eq5 
         Caption         =   "Copy &5"
      End
      Begin VB.Menu eq6 
         Caption         =   "Copy &6"
      End
      Begin VB.Menu eq7 
         Caption         =   "Copy &7"
      End
      Begin VB.Menu eq8 
         Caption         =   "Copy &8"
      End
      Begin VB.Menu eq9 
         Caption         =   "Copy &9"
      End
      Begin VB.Menu eq10 
         Caption         =   "Copy 1&0"
      End
      Begin VB.Menu eq11 
         Caption         =   "Copy 11"
      End
      Begin VB.Menu eq12 
         Caption         =   "Copy 12"
      End
      Begin VB.Menu eq13 
         Caption         =   "Copy 13"
      End
   End
   Begin VB.Menu CapItem 
      Caption         =   "&Capture"
      Begin VB.Menu DoCap 
         Caption         =   "Ca&pture"
      End
      Begin VB.Menu ShowCap 
         Caption         =   "&Show Capture Window"
      End
   End
End
Attribute VB_Name = "GraphPlot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(8) As Integer
Dim i As Integer
Dim i1 As Integer


Private Sub aboutme_Click()
frmAbout.Show
End Sub

Private Sub Command1_Click()
On Error Resume Next
Dim x As Double, y As Double
Dim delta As Double
Dim equ1 As String, equ2 As String
Dim j As Integer

With wait.ProgressBar1
    .Min = LBound(a)
    .Max = UBound(a)
    .Visible = True
    .Value = .Min
End With

Call RndColor
    Picture1.ForeColor = RGB(M, n, O)
    Text4.Text = Val(Text4.Text) - 5
    Picture2.Line (10, Val(Text4.Text))-(-10, Val(Text4.Text)), RGB(M, n, O)
   
For j = 0 To 8
    a(j) = Val(Text1(j).Text)
Next j

If a(0) = 0 And a(4) = 0 And a(5) = 0 And a(6) = 0 And a(1) <> 0 Then
    delta = a(2) ^ 2 - 4 * a(1) * a(3)
    If delta >= 0 Then
        Label19.Visible = False
        Label12.Visible = True
        Text2.Visible = True
        Text2.Text = (-a(2) + Sqr(delta)) / (2 * a(1))
        Text3.Visible = True
        Text3.Text = (-a(2) - Sqr(delta)) / (2 * a(1))
    Else
        Label12.Visible = False
        Text2.Visible = False
        Text3.Visible = False
        Label19.Visible = True
        End If
    Else
    Text2.Visible = False
    Text3.Visible = False
    Label12.Visible = False
    Label19.Visible = False
End If

wait.Show


StatusBar1.Panels(3).Text = "Please wait while drawing......"
For x = -a(8) To a(8) Step 0.001
    y = (a(0) * x ^ 3 + a(1) * x ^ 2 + a(2) * x + a(3)) / (a(4) * x ^ 3 + a(5) * x ^ 2 + a(6) * x + a(7))
    Picture1.PSet (x, y) 'draws the graph
    wait.ProgressBar1.Value = x
Next x
wait.Hide
StatusBar1.Panels(3).Text = ""

With wait.ProgressBar1
    .Visible = False
    .Value = .Min
End With

equ1 = "y= (" & a(0) & "x^3+" & a(1) & "x^2+" & a(2) & "x+" & a(3) & ") / (" & _
    a(4) & "x^3+" & a(5) & "x^2+" & a(6) & "x+" & a(7) & ")"

equ2 = "y = " & a(0) & "x^3+" & a(1) & "x^2+" & a(2) & "x+" & a(3)


If Val(Text5.Text) <= 12 Then
    
    Label1(Text5.Text).ForeColor = RGB(M, n, O)
    If a(0) = 0 And a(1) = 0 And a(2) = 0 And a(3) = 0 Then
        Label1(Text5.Text).Caption = "y = 0"
    ElseIf a(4) = 0 And a(5) = 0 And a(6) = 0 And a(7) = 1 Then
        Label1(Text5.Text).Caption = equ2
    Else
        Label1(Text5.Text).Caption = equ1
    End If
    Text5.Text = Val(Text5.Text) + 1
Else
    
    Call sort
    GraphPlot.Label1(12).ForeColor = RGB(M, n, O)
    If a(4) = 0 And a(5) = 0 And a(6) = 0 And a(7) = 1 Then
        Label1(12).Caption = equ2
    Else
        Label1(12).Caption = equ1
    End If
End If
End Sub

Private Sub Reset()
FClear = True
Text1(8).Text = 10
Picture1.Scale (-10, 10)-(10, -10)
Picture1.Cls 'clear picture box
Picture2.Cls

Text4.Text = 103 'reset the index
Text5.Text = 0

For i = 0 To 12
    Label1(i).Caption = ""
Next i

For i = 0 To 6
    Text1(i).Text = ""
Next i

For i = 0 To 12
    Line2(i).Visible = False
Next i

Text1(7) = 1
Text2.Visible = False
Text3.Visible = False
Label12.Visible = False
Label19.Visible = False

Picture1.Line (0, -10)-(0, 10), QBColor(0) 'Draw y-axis
Picture1.Line (-10, 0)-(10, 0), QBColor(0) 'Draw x-axis
For i1 = -10 To 10
    GraphPlot.Picture1.Line (-0.1, i1)-(0.1, i1), QBColor(0) 'Draw index
    GraphPlot.Picture1.Line (i1, -0.1)-(i1, 0.1), QBColor(0)
Next i1

End Sub


Private Sub Command2_Click()
Call Reset
End Sub

Private Sub eq11_Click()
Clipboard.Clear
Clipboard.SetText Label1(10).Caption
End Sub
Private Sub eq12_Click()
Clipboard.Clear
Clipboard.SetText Label1(11).Caption
End Sub

Private Sub eq13_Click()
Clipboard.Clear
Clipboard.SetText Label1(13).Caption
End Sub

Private Sub exitItem_Click()
GraphPlot.Hide
calc7.Show
End Sub
Private Sub Form_Activate()
Picture1.Scale (-a(8), a(8))-(a(8), -a(8))   'Scale
Picture2.Scale (-15, 100)-(15, -100)
Picture1.Line (0, -a(8))-(0, a(8)), QBColor(0) 'Draw y-axis
Picture1.Line (-a(8), 0)-(a(8), 0), QBColor(0) 'Draw x-axis
For i1 = -a(8) To a(8)
        Picture1.Line (-0.1, i1)-(0.1, i1), QBColor(0)
        Picture1.Line (i1, -0.1)-(i1, 0.1), QBColor(0)
Next i1
End Sub
Private Sub Form_Initialize()
Call InitCommonControls
End Sub
Private Sub Form_Load()
Randomize
Load wait
Ftrigo = False
Me.WindowState = 2
a(8) = 10
wait.ProgressBar1.Align = 2

Text2.Visible = False
Text3.Visible = False

Label12.Visible = False

Picture1.BackColor = RGB(246, 254, 254)
Picture1.AutoRedraw = True
Picture2.BackColor = RGB(250, 250, 228)
Picture2.AutoRedraw = True

For i = 0 To 2
    Text1(i).ToolTipText = "Enter the value of coefficient here"
Next i
Text1(3).ToolTipText = "Enter a constant here"
Text1(7).ToolTipText = "Enter a constant here"
For i = 4 To 6
    Text1(i).ToolTipText = "Enter the value of coefficient here"
Next i

For i = 0 To 12
    Line2(i).Visible = False
Next i
StatusBar1.Panels(3).Bevel = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    StatusBar1.Panels(6).Text = ""
End Sub

Private Sub item10_Click()
Text1(8).Text = 10
Picture1.Cls
Picture1.Scale (-10, 10)-(10, -10)   'Scale
Picture1.Line (0, -10)-(0, 10), QBColor(0) 'Draw y-axis
Picture1.Line (-10, 0)-(10, 0), QBColor(0) 'Draw x-axis
For i1 = -10 To 10
        Picture1.Line (-0.1, i1)-(0.1, i1), QBColor(0)
        Picture1.Line (i1, -0.1)-(i1, 0.1), QBColor(0)
Next i1
End Sub
Private Sub item100_Click()
Text1(8).Text = 100
Picture1.Cls
Picture1.Scale (-100, 100)-(100, -100)
Picture1.Line (0, -100)-(0, 100), QBColor(0) 'Draw y-axis
Picture1.Line (-100, 0)-(100, 0), QBColor(0) 'Draw x-axis
For i1 = -100 To 100 Step 10
    Picture1.Line (-1, i1)-(1, i1), QBColor(0)
    Picture1.Line (i1, -1)-(i1, 1), QBColor(0)
Next i1
End Sub
Private Sub item20_Click()
Text1(8).Text = 20
Picture1.Cls
Picture1.Scale (-20, 20)-(20, -20)
Picture1.Line (0, -20)-(0, 20), QBColor(0) 'Draw y-axis
Picture1.Line (-20, 0)-(20, 0), QBColor(0) 'Draw x-axis
For i1 = -20 To 20
    Picture1.Line (-0.1, i1)-(0.1, i1), QBColor(0)
    Picture1.Line (i1, -0.1)-(i1, 0.1), QBColor(0)
Next i1
End Sub
Private Sub item30_Click()
Text1(8).Text = 30
Picture1.Cls
Picture1.Scale (-30, 30)-(30, -30)
Picture1.Line (0, -30)-(0, 30), QBColor(0) 'Draw y-axis
Picture1.Line (-30, 0)-(30, 0), QBColor(0) 'Draw x-axis
For i1 = -30 To 30
    Picture1.Line (-0.3, i1)-(0.3, i1), QBColor(0)
    Picture1.Line (i1, -0.3)-(i1, 0.3), QBColor(0)
Next i1
End Sub
Private Sub item40_Click()
Text1(8).Text = 40
Picture1.Cls
Picture1.Scale (-40, 40)-(40, -40)
Picture1.Line (0, -40)-(0, 40), QBColor(0) 'Draw y-axis
Picture1.Line (-40, 0)-(40, 0), QBColor(0) 'Draw x-axis
For i1 = -40 To 40 Step 2
    Picture1.Line (-0.3, i1)-(0.3, i1), QBColor(0)
    Picture1.Line (i1, -0.3)-(i1, 0.3), QBColor(0)
Next i1
End Sub
Private Sub item50_Click()
Text1(8).Text = 50
Picture1.Cls
Picture1.Scale (-50, 50)-(50, -50)
Picture1.Line (0, -50)-(0, 50), QBColor(0) 'Draw y-axis
Picture1.Line (-50, 0)-(50, 0), QBColor(0) 'Draw x-axis
For i1 = -50 To 50 Step 2
    Picture1.Line (-0.3, i1)-(0.1, i1), QBColor(0)
    Picture1.Line (i1, -0.3)-(i1, 0.3), QBColor(0)
Next i1
End Sub
Private Sub item60_Click()
Text1(8).Text = 60
Picture1.Cls
Picture1.Scale (-60, 60)-(60, -60)
Picture1.Line (0, -60)-(0, 60), QBColor(0) 'Draw y-axis
Picture1.Line (-60, 0)-(60, 0), QBColor(0) 'Draw x-axis
For i1 = -60 To 60 Step 2
    Picture1.Line (-0.3, i1)-(0.3, i1), QBColor(0)
    Picture1.Line (i1, -0.3)-(i1, 0.3), QBColor(0)
Next i1
End Sub
Private Sub item70_Click()
Text1(8).Text = 70
Picture1.Cls
Picture1.Scale (-70, 70)-(70, -70)
Picture1.Line (0, -70)-(0, 70), QBColor(0) 'Draw y-axis
Picture1.Line (-70, 0)-(70, 0), QBColor(0) 'Draw x-axis
For i1 = -70 To 70 Step 5
    Picture1.Line (-0.5, i1)-(0.5, i1), QBColor(0)
    Picture1.Line (i1, -0.5)-(i1, 0.5), QBColor(0)
Next i1
End Sub
Private Sub item80_Click()
Text1(8).Text = 80
Picture1.Cls
Picture1.Scale (-80, 80)-(80, -80)
Picture1.Line (0, -80)-(0, 80), QBColor(0) 'Draw y-axis
Picture1.Line (-80, 0)-(80, 0), QBColor(0) 'Draw x-axis
For i1 = -80 To 80 Step 5
    Picture1.Line (-0.5, i1)-(0.5, i1), QBColor(0)
    Picture1.Line (i1, -0.5)-(i1, 0.5), QBColor(0)
Next i1
End Sub
Private Sub item90_Click()
Text1(8).Text = 90
Picture1.Cls
Picture1.Scale (-90, 90)-(90, -90)
Picture1.Line (0, -90)-(0, 90), QBColor(0) 'Draw y-axis
Picture1.Line (-90, 0)-(90, 0), QBColor(0) 'Draw x-axis
For i1 = -90 To 90 Step 5
    Picture1.Line (-1, i1)-(1, i1), QBColor(0)
    Picture1.Line (i1, -1)-(i1, 1), QBColor(0)
Next i1
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Select Case Index
Case 0
    StatusBar1.Panels(6).Text = "History 1"
Case 1
    StatusBar1.Panels(6).Text = "History 2"
Case 2
    StatusBar1.Panels(6).Text = "History 3"
Case 3
    StatusBar1.Panels(6).Text = "History 4"
Case 4
    StatusBar1.Panels(6).Text = "History 5"
Case 5
    StatusBar1.Panels(6).Text = "History 6"
Case 6
    StatusBar1.Panels(6).Text = "History 7"
Case 7
    StatusBar1.Panels(6).Text = "History 8"
Case 8
    StatusBar1.Panels(6).Text = "History 9"
Case 9
    StatusBar1.Panels(6).Text = "History 10"
Case 10
    StatusBar1.Panels(6).Text = "History 11"
Case 11
    StatusBar1.Panels(6).Text = "History 12"
Case 12
    StatusBar1.Panels(6).Text = "History 13"
End Select

   
End Sub

Private Sub Label5_Click()
If Ftrigo = False Then
Else
MsgBox "The coordinate in blue color is in Degree.", vbInformation, "Simple Graphic Plotter"
End If
End Sub
Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'this shows the cursor position on the picture box
If Ftrigo = True Then
    Label5.ForeColor = RGB(70, 1, 255)
    Label5 = Int((Int(x * 100) / 100) * (180 / 3.14))
Else
    Label5.ForeColor = QBColor(0)
    Label5 = Int(x * 100) / 100
End If
Label3 = Int(y * 100) / 100
End Sub
Private Sub clearItem_Click()
Call Reset
End Sub
Private Sub eq1_Click()
Clipboard.Clear
Clipboard.SetText Label1(0).Caption
End Sub
Private Sub eq10_Click()
Clipboard.Clear
Clipboard.SetText Label1(9).Caption
End Sub
Private Sub eq2_Click()
Clipboard.Clear
Clipboard.SetText Label1(1).Caption
End Sub
Private Sub eq3_Click()
Clipboard.Clear
Clipboard.SetText Label1(2).Caption
End Sub
Private Sub eq4_Click()
Clipboard.Clear
Clipboard.SetText Label1(3).Caption
End Sub
Private Sub eq5_Click()
Clipboard.Clear
Clipboard.SetText Label1(4).Caption
End Sub
Private Sub eq6_Click()
Clipboard.Clear
Clipboard.SetText Label1(5).Caption
End Sub
Private Sub eq7_Click()
Clipboard.Clear
Clipboard.SetText Label1(6).Caption
End Sub
Private Sub eq8_Click()
Clipboard.Clear
Clipboard.SetText Label1(7).Caption
End Sub
Private Sub eq9_Click()
Clipboard.Clear
Clipboard.SetText Label1(8).Caption
End Sub


Private Sub triItem_Click()
Ftrigo = True
Trigo.Show
End Sub
