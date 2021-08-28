VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{A40B4820-D5FD-11D1-8818-C199198E9702}#1.8#0"; "MMTOOLSX.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "gaming department of 1990"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MousePointer    =   2  'Cross
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MMToolsX.MMGaugeX MMGaugeX1 
      Height          =   1455
      Left            =   0
      TabIndex        =   12
      Top             =   7320
      Width           =   15135
      ForeColor       =   4210688
      MaxValue        =   1000
      Progress        =   1000
      _Handle         =   1
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "game.frx":0000
      Height          =   2895
      Left            =   9720
      TabIndex        =   32
      Top             =   1920
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5106
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "score"
         Caption         =   "score"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "lives left"
         Caption         =   "lives left"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "level"
         Caption         =   "level"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "name"
         Caption         =   "name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text5 
      Height          =   615
      Left            =   11400
      TabIndex        =   31
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton Command10 
      Caption         =   "search"
      Height          =   615
      Left            =   9600
      TabIndex        =   30
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton Command9 
      Caption         =   "empty text"
      Height          =   495
      Left            =   4080
      TabIndex        =   29
      Top             =   10080
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      Caption         =   "enable"
      Height          =   495
      Left            =   3960
      TabIndex        =   28
      Top             =   9480
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "disable"
      Height          =   495
      Left            =   3960
      TabIndex        =   27
      Top             =   8880
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   -240
      Picture         =   "game.frx":0015
      ScaleHeight     =   675
      ScaleWidth      =   4035
      TabIndex        =   26
      Top             =   3360
      Visible         =   0   'False
      Width           =   4095
      Begin VB.Timer Timer11 
         Interval        =   1
         Left            =   360
         Top             =   120
      End
   End
   Begin VB.Timer Timer10 
      Interval        =   1
      Left            =   7560
      Top             =   3480
   End
   Begin VB.Timer Timer9 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9360
      Top             =   8760
   End
   Begin VB.TextBox Text3 
      DataField       =   "level"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   7800
      TabIndex        =   21
      Top             =   10440
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      DataField       =   "lives left"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   7800
      TabIndex        =   20
      Top             =   9840
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      DataField       =   "score"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   7800
      TabIndex        =   19
      Top             =   9240
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "delete"
      Height          =   375
      Left            =   6720
      TabIndex        =   18
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      DataField       =   "name"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   615
      Left            =   6720
      TabIndex        =   17
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "add new"
      Height          =   375
      Left            =   6720
      TabIndex        =   16
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "refresh/restart"
      Height          =   375
      Left            =   5760
      TabIndex        =   15
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "save"
      Height          =   735
      Left            =   9720
      TabIndex        =   14
      Top             =   9240
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   240
      Negotiate       =   -1  'True
      Top             =   10200
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1085
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
      BackColor       =   -2147483624
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"game.frx":1074
      OLEDBString     =   $"game.frx":1136
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "game"
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
   Begin VB.Timer Timer8 
      Interval        =   1
      Left            =   4440
      Top             =   6240
   End
   Begin VB.Timer Timer7 
      Interval        =   1
      Left            =   4440
      Top             =   5760
   End
   Begin VB.Timer Timer6 
      Interval        =   1
      Left            =   4440
      Top             =   5280
   End
   Begin VB.CommandButton Command1 
      Caption         =   "start"
      Enabled         =   0   'False
      Height          =   735
      Left            =   5760
      TabIndex        =   7
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton shape2 
      Caption         =   "bomb"
      BeginProperty Font 
         Name            =   "HL2cross"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton shape1 
      Caption         =   "bomb"
      BeginProperty Font 
         Name            =   "HL2cross"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Shape4 
      Caption         =   "bomb"
      BeginProperty Font 
         Name            =   "HL2cross"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      TabIndex        =   4
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Shape3 
      Caption         =   "bomb"
      BeginProperty Font 
         Name            =   "HL2cross"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      TabIndex        =   3
      Top             =   2400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer5 
      Interval        =   1
      Left            =   4440
      Top             =   4800
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   2160
      Top             =   3120
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   2160
      Top             =   2400
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   2160
      Top             =   1560
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   2160
      Top             =   960
   End
   Begin VB.Label Label12 
      Caption         =   "enter your name here"
      Height          =   495
      Left            =   4560
      TabIndex        =   25
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label11 
      Caption         =   "level"
      Height          =   375
      Left            =   6480
      TabIndex        =   24
      Top             =   10560
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "lives left"
      Height          =   495
      Left            =   6480
      TabIndex        =   23
      Top             =   9840
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "score"
      Height          =   375
      Left            =   6480
      TabIndex        =   22
      Top             =   9240
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "people left in percentage"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   975
      Left            =   4920
      TabIndex        =   13
      Top             =   6240
      Width           =   7815
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
      Left            =   120
      TabIndex        =   10
      Top             =   6120
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
      Left            =   360
      TabIndex        =   9
      Top             =   5040
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
      Left            =   600
      TabIndex        =   8
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "1"
      DataField       =   "level"
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
      Left            =   1560
      TabIndex        =   2
      Top             =   6120
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "1000"
      DataField       =   "lives left"
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
      Left            =   1800
      TabIndex        =   1
      Top             =   5160
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      DataField       =   "score"
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
      Left            =   1800
      TabIndex        =   0
      Top             =   4080
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
End Sub

Private Sub Command10_Click()
If Text5.Text = "" Then
MsgBox "enter value for searching a record"

Else

    str1 = Text5.Text
    strsearch = "name like '" & str1 & "'"
    Adodc1.Recordset.MoveFirst
    Adodc1.Recordset.Find (strsearch)
    MsgBox "results found"
End If
End Sub

Private Sub Command2_Click()
On Error GoTo err
Adodc1.Recordset.Save
Exit Sub
err:
End Sub

Private Sub Command3_Click()
Unload Form1
Load Form1
Form1.Show
End Sub

Private Sub Command4_Click()
On Error GoTo err
Adodc1.Recordset.AddNew
Text4.Enabled = True
Exit Sub
err:
End Sub

Private Sub Command5_Click()
On Error GoTo err
Adodc1.Recordset.Delete
Exit Sub
err:
End Sub

Private Sub Command7_Click()
Timer9.Enabled = False
End Sub

Private Sub Command8_Click()
Timer9.Enabled = True
End Sub

Private Sub Command9_Click()
Text1.Text = Empty
Text2.Text = Empty
Text3.Text = Empty
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

Private Sub shape1_Click()
MsgBox "too high"
End Sub

Private Sub shape2_Click()
MsgBox " poison removed"
Label1.Caption = Val(Label1.Caption) + 10
Label2.Caption = Val(Label2.Caption) + 75
MMGaugeX1.Progress = Val(MMGaugeX1.Progress) + 75
End Sub

Private Sub Shape3_Click()
MsgBox "poison removed!"
Label1.Caption = Val(Label1.Caption) + 10
Label2.Caption = Val(Label2.Caption) + 75
MMGaugeX1.Progress = Val(MMGaugeX1.Progress) + 75
End Sub

Private Sub Text4_Change()
If Text4.Text = Empty Then
Command1.Enabled = False
Else
Command1.Enabled = True
End If
End Sub

Private Sub Timer1_Timer()
shape1.Visible = True
Timer2.Enabled = True
Shape4.Visible = False
Timer1.Enabled = False
Timer5.Enabled = True
End Sub

Private Sub Timer10_Timer()
Text4.Text = Empty
Timer10.Enabled = False
End Sub

Private Sub Timer11_Timer()
If Shape4.Visible = True Then
Picture1.Visible = True
Else
Picture1.Visible = False
End If
End Sub

Private Sub Timer12_Timer()

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
MMGaugeX1.Progress = Val(MMGaugeX1.Progress) - 75
End If
End Sub

Private Sub Timer6_Timer()
If Label1.Caption = 100 Then
Timer1.Interval = Val(Timer1.Interval) - 80
Timer2.Interval = Val(Timer2.Interval) - 80
Timer3.Interval = Val(Timer3.Interval) - 80
Timer4.Interval = Val(Timer4.Interval) - 80
Timer6.Enabled = False
MsgBox "next level"
Label3.Caption = 2
End If
End Sub

Private Sub Timer7_Timer()
If Label3.Caption = 3 Then
End If
End Sub

Private Sub Timer8_Timer()
If Label1.Caption = 200 Then
Label3.Caption = 3
Timer8.Enabled = False
End If
End Sub

Private Sub Timer9_Timer()
Text1.Text = Label1.Caption
Text2.Text = Label2.Caption
Text3.Text = Label3.Caption
End Sub
