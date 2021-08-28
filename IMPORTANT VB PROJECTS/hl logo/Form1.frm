VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000012&
   Caption         =   "Form1"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Form1.frx":0000
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "show form5"
      Height          =   855
      Left            =   4560
      TabIndex        =   8
      Top             =   4800
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "show form3"
      Height          =   855
      Left            =   2880
      TabIndex        =   7
      Top             =   4800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "show form4"
      Height          =   855
      Left            =   1200
      TabIndex        =   6
      Top             =   4800
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Timer Timer19 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1080
      Top             =   120
   End
   Begin VB.Timer Timer18 
      Interval        =   250
      Left            =   2400
      Top             =   1080
   End
   Begin VB.Timer Timer17 
      Interval        =   3500
      Left            =   3840
      Top             =   1080
   End
   Begin VB.Timer Timer16 
      Interval        =   3000
      Left            =   3480
      Top             =   1080
   End
   Begin VB.Timer Timer15 
      Interval        =   2500
      Left            =   3120
      Top             =   1080
   End
   Begin VB.Timer Timer14 
      Interval        =   500
      Left            =   2760
      Top             =   1080
   End
   Begin VB.Timer Timer13 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   4200
      Top             =   1080
   End
   Begin VB.Timer Timer12 
      Interval        =   1
      Left            =   3480
      Top             =   720
   End
   Begin VB.Timer Timer11 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   3120
      Top             =   720
   End
   Begin VB.Timer Timer10 
      Interval        =   1
      Left            =   2760
      Top             =   720
   End
   Begin VB.Timer Timer9 
      Interval        =   50
      Left            =   2400
      Top             =   720
   End
   Begin VB.Timer Timer8 
      Interval        =   1
      Left            =   3480
      Top             =   360
   End
   Begin VB.Timer Timer7 
      Interval        =   20
      Left            =   3120
      Top             =   360
   End
   Begin VB.Timer Timer6 
      Interval        =   1
      Left            =   2760
      Top             =   360
   End
   Begin VB.Timer Timer5 
      Interval        =   20
      Left            =   2400
      Top             =   360
   End
   Begin VB.Timer Timer4 
      Interval        =   1
      Left            =   3480
      Top             =   0
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   3120
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   2760
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   2400
      Top             =   0
   End
   Begin VB.Label Label5 
      BackColor       =   &H000000FF&
      Caption         =   "click on"" Enter"" to proceed to next page"
      Height          =   855
      Left            =   -1480
      TabIndex        =   5
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "A    c    a    d    e    m    y"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   495
      Left            =   2160
      TabIndex        =   3
      Top             =   2400
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   24
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   495
      Left            =   3480
      MouseIcon       =   "Form1.frx":1272
      TabIndex        =   2
      Top             =   3240
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   2295
      Left            =   600
      Top             =   1560
      Visible         =   0   'False
      Width           =   6975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Test"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   240
      Left            =   3555
      TabIndex        =   1
      Top             =   1680
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "S a i n t m a r y ' s"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   2475
      Left            =   1440
      TabIndex        =   0
      Top             =   1680
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   5535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Image1_Click()

End Sub

Private Sub Command2_Click()
Form1.Hide
form3.Show
End Sub

Private Sub Command1_Click()
Form1.Hide
Form4.Show
End Sub



Private Sub Command3_Click()
Form1.Hide
Form5.Show
End Sub

Private Sub Form_Load()
MsgBox "important:when you make any changes in the program then immediatly restart the program to apply them and avoid a program crash "
End Sub

Private Sub Label4_Click()
Timer19.Enabled = True
End Sub

Private Sub Label5_Click()
Label5.Left = -1480
End Sub

Private Sub Label6_Click()
Form1.Hide
Form2.Show
End Sub

Private Sub Timer1_Timer()
Label1.Left = Val(Label1.Left) - 10
End Sub

Private Sub Timer10_Timer()
If Label3.Font.Size >= 70 Then
Timer11.Enabled = True
Timer9.Enabled = False
End If
End Sub

Private Sub Timer11_Timer()
Label3.Font.Size = Val(Label3.Font.Size) - 1
End Sub

Private Sub Timer12_Timer()
If Label3.Font.Size < 50 Then
Timer11.Enabled = False
Timer9.Enabled = True
End If
End Sub

Private Sub Timer13_Timer()
Command2.Visible = True
Form1.Hide
Form2.Show
End Sub

Private Sub Timer14_Timer()
Label1.Visible = True
End Sub

Private Sub Timer15_Timer()
Label6.Visible = True
End Sub

Private Sub Timer16_Timer()
Label3.Visible = True
End Sub

Private Sub Timer17_Timer()
Label2.Visible = True
End Sub

Private Sub Timer18_Timer()
Shape1.Visible = True
End Sub

Private Sub Timer19_Timer()
Label5.Left = Val(Label5.Left) + 50
If Label5.Left >= 0 Then
Timer19.Enabled = False
End If
End Sub

Private Sub Timer2_Timer()
If Label1.Left <= 900 Then Timer3.Enabled = True
End Sub

Private Sub Timer3_Timer()
Timer1.Enabled = False
Label1.Left = Val(Label1.Left) + 10
End Sub

Private Sub Timer4_Timer()
If Label1.Left >= 2000 Then
Timer3.Enabled = False
Timer1.Enabled = True
End If
End Sub

Private Sub Timer5_Timer()
Label2.Left = Val(Label2.Left) - 10
End Sub

Private Sub Timer6_Timer()
If Label2.Left <= 1600 Then Timer7.Enabled = True
End Sub

Private Sub Timer7_Timer()
Timer5.Enabled = False
Label2.Left = Val(Label2.Left) + 10
End Sub

Private Sub Timer8_Timer()
If Label2.Left >= 2700 Then
Timer7.Enabled = False
Timer5.Enabled = True
End If
End Sub

Private Sub Timer9_Timer()
Label3.Font.Size = Val(Label3.Font.Size + 1)
End Sub
