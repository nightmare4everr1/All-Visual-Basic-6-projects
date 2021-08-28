VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form4 
   BackColor       =   &H80000008&
   Caption         =   "Form4"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8025
   LinkTopic       =   "Form4"
   ScaleHeight     =   8310
   ScaleWidth      =   8025
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   1080
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   1440
      Top             =   0
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   1800
      Top             =   0
   End
   Begin VB.Timer Timer4 
      Interval        =   1
      Left            =   2160
      Top             =   0
   End
   Begin VB.Timer Timer5 
      Interval        =   20
      Left            =   1080
      Top             =   360
   End
   Begin VB.Timer Timer6 
      Interval        =   1
      Left            =   1440
      Top             =   360
   End
   Begin VB.Timer Timer7 
      Interval        =   20
      Left            =   1800
      Top             =   360
   End
   Begin VB.Timer Timer8 
      Interval        =   1
      Left            =   2160
      Top             =   360
   End
   Begin VB.Timer Timer9 
      Interval        =   50
      Left            =   1080
      Top             =   720
   End
   Begin VB.Timer Timer10 
      Interval        =   1
      Left            =   1440
      Top             =   720
   End
   Begin VB.Timer Timer11 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1800
      Top             =   720
   End
   Begin VB.Timer Timer12 
      Interval        =   1
      Left            =   2160
      Top             =   720
   End
   Begin VB.Timer Timer13 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   2880
      Top             =   1080
   End
   Begin VB.Timer Timer14 
      Interval        =   500
      Left            =   1440
      Top             =   1080
   End
   Begin VB.Timer Timer15 
      Interval        =   2500
      Left            =   1800
      Top             =   1080
   End
   Begin VB.Timer Timer16 
      Interval        =   3000
      Left            =   2160
      Top             =   1080
   End
   Begin VB.Timer Timer17 
      Interval        =   3500
      Left            =   2520
      Top             =   1080
   End
   Begin VB.Timer Timer18 
      Interval        =   250
      Left            =   1080
      Top             =   1080
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   495
      Left            =   3480
      Top             =   7440
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483641
      ForeColor       =   -2147483639
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Microsoft Visual Studio\VB98\hl logo\quiz.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Microsoft Visual Studio\VB98\hl logo\quiz.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "score"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "create a print"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   5760
      TabIndex        =   13
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   585
      Left            =   5880
      Picture         =   "Form4.frx":0000
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   600
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
      Left            =   1680
      TabIndex        =   12
      Top             =   1680
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   5535
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
      Left            =   3795
      TabIndex        =   11
      Top             =   1680
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   2295
      Left            =   840
      Top             =   1560
      Visible         =   0   'False
      Width           =   6975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
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
      Left            =   3720
      MouseIcon       =   "Form4.frx":066A
      TabIndex        =   10
      Top             =   3240
      Visible         =   0   'False
      Width           =   855
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
      Left            =   2400
      TabIndex        =   9
      Top             =   2400
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Label9"
      DataField       =   "totalmarks"
      DataSource      =   "Adodc2"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   4320
      TabIndex        =   8
      Top             =   6840
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      DataField       =   "canidatescore"
      DataSource      =   "Adodc2"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Label7"
      DataField       =   "canidaterol no"
      DataSource      =   "Adodc2"
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   4320
      TabIndex        =   6
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Label5"
      DataField       =   "canidatename"
      DataSource      =   "Adodc2"
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   4320
      TabIndex        =   5
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000012&
      Caption         =   "your score"
      BeginProperty Font 
         Name            =   "HL2cross"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   1215
      Left            =   480
      TabIndex        =   4
      Top             =   5160
      Width           =   2295
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "student name"
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   5160
      Width           =   855
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "roll no"
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "canidate score"
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   6240
      Width           =   975
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "max marks"
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   3000
      TabIndex        =   0
      Top             =   6840
      Width           =   975
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Adodc2.Recordset.MoveLast
End Sub

Private Sub Image1_Click()
Form4.Hide
Form5.Show
End Sub

Private Sub Label1_Click()
End
End Sub

Private Sub Label4_Click()
Dim userMsg As String
userMsg = InputBox("input password", "password", "Enter your password here", 500, 700)
If userMsg = "kai" Then
MsgBox "correct match reproducing records..."
Adodc2.Visible = True
Adodc2.Enabled = True
Else
MsgBox "error!"
End If
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

