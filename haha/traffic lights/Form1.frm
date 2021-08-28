VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form Form1 
   Caption         =   "traffic lights"
   ClientHeight    =   9690
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11400
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MouseIcon       =   "Form1.frx":0000
   ScaleHeight     =   9690
   ScaleWidth      =   11400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "change timers"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      TabIndex        =   17
      Top             =   6240
      Width           =   2895
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Text            =   $"Form1.frx":0152
      Top             =   8640
      Width           =   9375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "set timer values"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      MouseIcon       =   "Form1.frx":01E8
      TabIndex        =   14
      Top             =   5400
      Width           =   2895
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   7320
      Top             =   8040
   End
   Begin VB.CommandButton Command5 
      Caption         =   "start traffic lights"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4320
      TabIndex        =   13
      Top             =   7200
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "reset traffic lights to original position"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6240
      TabIndex        =   12
      Top             =   7200
      Width           =   1575
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2175
      Left            =   4320
      TabIndex        =   5
      Top             =   4920
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   3836
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "red light time"
      TabPicture(0)   =   "Form1.frx":033A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Text1"
      Tab(0).Control(1)=   "Label7"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "yellow light time"
      TabPicture(1)   =   "Form1.frx":0356
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text2"
      Tab(1).Control(1)=   "Label6"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "green light time"
      TabPicture(2)   =   "Form1.frx":0372
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label8"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Text3"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1440
         TabIndex        =   8
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -73440
         TabIndex        =   7
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -73320
         TabIndex        =   6
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "seconds"
         Height          =   375
         Left            =   1560
         TabIndex        =   11
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "seconds"
         Height          =   375
         Left            =   -73320
         TabIndex        =   10
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "seconds"
         Height          =   375
         Left            =   -73200
         TabIndex        =   9
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.Timer Timer5 
      Interval        =   1000
      Left            =   360
      Top             =   240
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   5520
      Top             =   4440
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5520
      Top             =   3720
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   5520
      Top             =   3000
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5520
      Top             =   2400
   End
   Begin VB.Label Label9 
      Caption         =   "read these instructions before entering timer values"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   16
      Top             =   7920
      Width           =   3615
   End
   Begin VB.Label Label5 
      Caption         =   "restarts timers"
      Height          =   375
      Left            =   6120
      TabIndex        =   4
      Top             =   2400
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4920
      TabIndex        =   3
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "controls green light delay"
      Height          =   375
      Left            =   6000
      TabIndex        =   2
      Top             =   4440
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "controls yellow light delay"
      Height          =   375
      Left            =   6120
      TabIndex        =   1
      Top             =   3720
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "controls red light delay"
      Height          =   375
      Left            =   6000
      TabIndex        =   0
      Top             =   3000
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Shape Shape6 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   2055
      Left            =   2640
      Shape           =   3  'Circle
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   2055
      Left            =   2640
      Shape           =   3  'Circle
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Shape Shape4 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   2055
      Left            =   2640
      Shape           =   3  'Circle
      Top             =   480
      Width           =   1215
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   2055
      Left            =   2640
      Shape           =   3  'Circle
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   2055
      Left            =   2640
      Shape           =   3  'Circle
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   2055
      Left            =   2640
      Shape           =   3  'Circle
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
SSTab1.Enabled = True
MsgBox "please type in the password"
Form1.Hide
Form2.Show
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command4_Click()
Timer6.Enabled = True
End Sub

Private Sub Command5_Click()
Timer1.Enabled = True
End Sub

Private Sub Command6_Click()
Timer2.Interval = Text1.Text * 1000
Timer3.Interval = Text2.Text * 1000
Timer4.Interval = Text3.Text * 1000
End Sub

Private Sub Form_Load()
MsgBox "please verify yourself"
Form1.Hide
Form3.Show
End Sub

Private Sub Timer1_Timer()
Shape4.Visible = False
Timer2.Enabled = True
Timer1.Enabled = False
Shape6.Visible = True
End Sub

Private Sub Timer2_Timer()
Shape5.Visible = False
Timer3.Enabled = True
Timer2.Enabled = False
End Sub

Private Sub Timer3_Timer()
Shape4.Visible = True
Shape5.Visible = True
Shape6.Visible = False
Timer3.Enabled = False
Timer4.Enabled = True
End Sub

Private Sub Timer4_Timer()
Timer1.Enabled = True
Timer4.Enabled = False
End Sub

Private Sub Timer5_Timer()
Label4.Caption = Time
End Sub

Private Sub Timer6_Timer()
Shape4.Visible = True
Shape5.Visible = True
Shape6.Visible = True
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer6.Enabled = False
End Sub
