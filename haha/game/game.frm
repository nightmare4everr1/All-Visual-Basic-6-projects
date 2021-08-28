VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MousePointer    =   2  'Cross
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   480
      Top             =   10440
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Microsoft Visual Studio\VB98\game.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Microsoft Visual Studio\VB98\game.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
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
   Begin VB.CommandButton Command1 
      Caption         =   "start"
      Height          =   735
      Left            =   6000
      TabIndex        =   7
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton shape2 
      Caption         =   "bomb"
      Height          =   735
      Left            =   840
      TabIndex        =   6
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton shape1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   840
      TabIndex        =   5
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton Shape4 
      Caption         =   "Command4"
      Height          =   615
      Left            =   840
      TabIndex        =   4
      Top             =   3360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Shape3 
      Caption         =   "Command3"
      Height          =   735
      Left            =   840
      TabIndex        =   3
      Top             =   2400
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Timer Timer5 
      Interval        =   1
      Left            =   6480
      Top             =   4920
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   1560
      Top             =   5520
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   1080
      Top             =   5520
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   600
      Top             =   5520
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   120
      Top             =   5520
   End
   Begin VB.Label Label7 
      Caption         =   "surface"
      Height          =   495
      Left            =   5400
      TabIndex        =   11
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   5280
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Label Label6 
      Caption         =   "level"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1920
      TabIndex        =   10
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "people still alive"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1920
      TabIndex        =   9
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "score"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      TabIndex        =   8
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3360
      TabIndex        =   2
      Top             =   6720
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "1000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3480
      TabIndex        =   1
      Top             =   5640
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "0"
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
      Left            =   3600
      TabIndex        =   0
      Top             =   4680
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Image1_Click()

End Sub

Private Sub Command1_Click()
Timer1.Enabled = True
OLE1.Enabled
End Sub

Private Sub Label2_Change()
If Label2.Caption <= 0 Then
MsgBox "you lose"
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Label2.Caption = 0
End If
End Sub

Private Sub Label3_Click()

End Sub

Private Sub Label6_Click()

End Sub

Private Sub shape1_Click()
MsgBox "too high"
End Sub

Private Sub Shape3_Click()
MsgBox " perfect shot"
Label1.Caption = Val(Label1.Caption) + 10
End Sub

Private Sub Timer1_Timer()
shape1.Visible = True
Timer2.Enabled = True
Shape4.Visible = False
Timer1.Enabled = False
Timer5.Enabled = True
End Sub

Private Sub Timer2_Timer()
shape1.Visible = False
shape2.Visible = True
Timer2.Enabled = False
Timer3.Enabled = True
End Sub

Private Sub Timer3_Timer()
shape2.Visible = False
Shape3.Visible = True
Timer3.Enabled = False
Timer4.Enabled = True
End Sub

Private Sub Timer4_Timer()
Shape3.Visible = False
Shape4.Visible = True
Timer4.Enabled = False
Timer1.Enabled = True
End Sub

Private Sub Timer5_Timer()
If Shape4.Visible = True Then
Label2.Caption = Val(Label2.Caption) - 75
Timer5.Enabled = False
End If
End Sub
