VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text13 
      Height          =   495
      Left            =   12360
      TabIndex        =   31
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "search"
      Height          =   615
      Left            =   10080
      TabIndex        =   30
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   6720
      Top             =   360
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "badminton.frx":0000
      Height          =   4095
      Left            =   0
      TabIndex        =   29
      Top             =   6000
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   7223
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   33
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
      ColumnCount     =   12
      BeginProperty Column00 
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
      BeginProperty Column01 
         DataField       =   "forehand"
         Caption         =   "forehand"
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
         DataField       =   "backhand"
         Caption         =   "backhand"
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
         DataField       =   "net shots"
         Caption         =   "net shots"
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
      BeginProperty Column04 
         DataField       =   "what is ur avg rally"
         Caption         =   "what is ur avg rally"
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
      BeginProperty Column05 
         DataField       =   "stamina"
         Caption         =   "stamina"
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
      BeginProperty Column06 
         DataField       =   "serve"
         Caption         =   "serve"
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
      BeginProperty Column07 
         DataField       =   "reflexes"
         Caption         =   "reflexes"
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
      BeginProperty Column08 
         DataField       =   "aggresive, defense, medium"
         Caption         =   "aggresive, defense, medium"
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
      BeginProperty Column09 
         DataField       =   "specialise in"
         Caption         =   "specialise in"
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
      BeginProperty Column10 
         DataField       =   "your age"
         Caption         =   "your age"
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
      BeginProperty Column11 
         DataField       =   "percentage"
         Caption         =   "percentage"
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
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1319.811
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1995.024
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   915.024
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5040
      Top             =   4920
   End
   Begin VB.TextBox Text12 
      DataField       =   "percentage"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1440
      TabIndex        =   26
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "refresh"
      Height          =   495
      Left            =   6240
      TabIndex        =   25
      ToolTipText     =   "use it if encountering a problem"
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "delete"
      Height          =   495
      Left            =   4680
      TabIndex        =   24
      ToolTipText     =   "permanently deletes one form"
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "add new"
      Height          =   495
      Left            =   6240
      TabIndex        =   23
      ToolTipText     =   "creates a blank form"
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "save"
      Height          =   495
      Left            =   4680
      TabIndex        =   22
      ToolTipText     =   "saves the data"
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox Text11 
      DataField       =   "your age"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3840
      TabIndex        =   10
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox Text10 
      DataField       =   "specialise in"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2280
      TabIndex        =   9
      Top             =   5400
      Width           =   3855
   End
   Begin VB.TextBox Text9 
      DataField       =   "aggresive, defense, medium"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2280
      TabIndex        =   8
      Top             =   4920
      Width           =   1455
   End
   Begin VB.TextBox Text8 
      DataField       =   "reflexes"
      DataSource      =   "Adodc1"
      Height          =   525
      Left            =   2280
      TabIndex        =   7
      Top             =   4440
      Width           =   1455
   End
   Begin VB.TextBox Text7 
      DataField       =   "serve"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2280
      TabIndex        =   6
      Top             =   3960
      Width           =   1455
   End
   Begin VB.TextBox Text6 
      DataField       =   "stamina"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2280
      TabIndex        =   5
      Top             =   3480
      Width           =   1455
   End
   Begin VB.TextBox Text5 
      DataField       =   "what is ur avg rally"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      DataField       =   "net shots"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      DataField       =   "backhand"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      DataField       =   "forehand"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   1560
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   4680
      Top             =   3600
      Width           =   3495
      _ExtentX        =   6165
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"badminton.frx":0015
      OLEDBString     =   $"badminton.frx":00A8
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "badminton"
      Caption         =   "scrolling bar"
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
   Begin VB.TextBox Text1 
      DataField       =   "name"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label13 
      Caption         =   "percentage"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   28
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label12 
      Caption         =   "total"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   27
      Top             =   -120
      Width           =   1575
   End
   Begin VB.Label Label11 
      Caption         =   "forehand"
      Height          =   375
      Left            =   1080
      TabIndex        =   21
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "backhand"
      Height          =   375
      Left            =   1080
      TabIndex        =   20
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "net shots"
      Height          =   375
      Left            =   1080
      TabIndex        =   19
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "what is your avg rally"
      Height          =   375
      Left            =   1080
      TabIndex        =   18
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "stamina"
      Height          =   375
      Left            =   1080
      TabIndex        =   17
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "serve"
      Height          =   375
      Left            =   1080
      TabIndex        =   16
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "reflexes"
      Height          =   375
      Left            =   1080
      TabIndex        =   15
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "aggresive,defensive,medium"
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "specialise in"
      Height          =   375
      Left            =   1080
      TabIndex        =   13
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "age"
      Height          =   375
      Left            =   3960
      TabIndex        =   12
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "name"
      Height          =   375
      Left            =   1080
      TabIndex        =   11
      Top             =   1080
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo err
Adodc1.Recordset.Save
Exit Sub
err:
Unload Form1
Load Form1
Form1.Show
End Sub

Private Sub Command2_Click()
On Error GoTo err
Adodc1.Recordset.AddNew
Exit Sub
err:
End Sub

Private Sub Command3_Click()
On Error GoTo err
Adodc1.Recordset.Delete
Exit Sub
err:
Unload Form1
Load Form1
Form1.Show
End Sub

Private Sub Command4_Click()
Unload Form1
Load Form1
Form1.Show
End Sub

Private Sub Command5_Click()
If Text13.Text = "" Then
MsgBox "enter value for searching a record"

Else

    str1 = Text13.Text
    strsearch = "name like '" & str1 & "'"

    Adodc1.Recordset.MoveFirst
    Adodc1.Recordset.Find (strsearch)
    MsgBox "search complete"
End If
End Sub

Private Sub Timer1_Timer()
Text12.Text = (Val(Text2.Text) + Val(Text3.Text) + Val(Text4.Text) + Val(Text6.Text) + Val(Text7.Text) + Val(Text8.Text)) * 100 / 72
End Sub

