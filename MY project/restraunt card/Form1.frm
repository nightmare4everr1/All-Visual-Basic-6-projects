VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "kfc service card"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13800
   LinkTopic       =   "Form1"
   ScaleHeight     =   11115
   ScaleWidth      =   13800
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   11400
      TabIndex        =   71
      Top             =   5040
      Width           =   1815
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form1.frx":0000
      Height          =   2175
      Left            =   8760
      TabIndex        =   70
      Top             =   5880
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   3836
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command6 
      Caption         =   "search"
      Height          =   375
      Left            =   9120
      TabIndex        =   69
      Top             =   5040
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      Caption         =   "clear checks"
      Height          =   495
      Left            =   9240
      TabIndex        =   68
      Top             =   3960
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "add new"
      Height          =   495
      Left            =   9000
      TabIndex        =   67
      Top             =   3360
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "refresh"
      Height          =   495
      Left            =   9000
      TabIndex        =   66
      Top             =   2760
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "delete"
      Height          =   615
      Left            =   9000
      TabIndex        =   65
      Top             =   2040
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "save"
      Height          =   495
      Left            =   9000
      TabIndex        =   64
      Top             =   1440
      Width           =   2775
   End
   Begin VB.CheckBox Check35 
      Caption         =   "Check5"
      DataField       =   "f5"
      DataSource      =   "Adodc1"
      Height          =   195
      Left            =   6960
      TabIndex        =   63
      Top             =   8040
      Width           =   255
   End
   Begin VB.CheckBox Check34 
      Caption         =   "Check4"
      DataField       =   "f4"
      DataSource      =   "Adodc1"
      Height          =   195
      Left            =   5880
      TabIndex        =   62
      Top             =   8040
      Width           =   255
   End
   Begin VB.CheckBox Check33 
      Caption         =   "Check3"
      DataField       =   "f3"
      DataSource      =   "Adodc1"
      Height          =   195
      Left            =   4680
      TabIndex        =   61
      Top             =   8040
      Width           =   255
   End
   Begin VB.CheckBox Check32 
      Caption         =   "Check2"
      DataField       =   "f2"
      DataSource      =   "Adodc1"
      Height          =   195
      Left            =   3480
      TabIndex        =   60
      Top             =   8040
      Width           =   255
   End
   Begin VB.CheckBox Check31 
      Caption         =   "Check1"
      DataField       =   "f1"
      DataSource      =   "Adodc1"
      Height          =   195
      Left            =   2400
      TabIndex        =   59
      Top             =   8040
      Width           =   255
   End
   Begin VB.CheckBox Check30 
      Caption         =   "Check5"
      DataField       =   "e5"
      DataSource      =   "Adodc1"
      Height          =   195
      Left            =   6960
      TabIndex        =   58
      Top             =   7320
      Width           =   255
   End
   Begin VB.CheckBox Check29 
      Caption         =   "Check4"
      DataField       =   "e4"
      DataSource      =   "Adodc1"
      Height          =   195
      Left            =   5880
      TabIndex        =   57
      Top             =   7320
      Width           =   255
   End
   Begin VB.CheckBox Check28 
      Caption         =   "Check3"
      DataField       =   "e3"
      DataSource      =   "Adodc1"
      Height          =   195
      Left            =   4680
      TabIndex        =   56
      Top             =   7320
      Width           =   255
   End
   Begin VB.CheckBox Check27 
      Caption         =   "Check2"
      DataField       =   "e2"
      DataSource      =   "Adodc1"
      Height          =   195
      Left            =   3480
      TabIndex        =   55
      Top             =   7320
      Width           =   255
   End
   Begin VB.CheckBox Check26 
      Caption         =   "Check1"
      DataField       =   "e1"
      DataSource      =   "Adodc1"
      Height          =   195
      Left            =   2400
      TabIndex        =   54
      Top             =   7320
      Width           =   255
   End
   Begin VB.CheckBox Check25 
      Caption         =   "Check5"
      DataField       =   "a5"
      DataSource      =   "Adodc1"
      Height          =   195
      Left            =   6960
      TabIndex        =   53
      Top             =   4320
      Width           =   255
   End
   Begin VB.CheckBox Check24 
      Caption         =   "Check4"
      DataField       =   "a4"
      DataSource      =   "Adodc1"
      Height          =   195
      Left            =   5880
      TabIndex        =   52
      Top             =   4320
      Width           =   255
   End
   Begin VB.CheckBox Check23 
      Caption         =   "Check3"
      DataField       =   "a3"
      DataSource      =   "Adodc1"
      Height          =   195
      Left            =   4680
      TabIndex        =   51
      Top             =   4320
      Width           =   255
   End
   Begin VB.CheckBox Check22 
      Caption         =   "Check2"
      DataField       =   "a2"
      DataSource      =   "Adodc1"
      Height          =   195
      Left            =   3480
      TabIndex        =   50
      Top             =   4320
      Width           =   255
   End
   Begin VB.CheckBox Check21 
      Caption         =   "Check1"
      DataField       =   "a 1"
      DataSource      =   "Adodc1"
      Height          =   195
      Left            =   2400
      TabIndex        =   49
      Top             =   4320
      Width           =   255
   End
   Begin VB.CheckBox Check20 
      Caption         =   "Check5"
      DataField       =   "b5"
      DataSource      =   "Adodc1"
      Height          =   195
      Left            =   6960
      TabIndex        =   48
      Top             =   5040
      Width           =   255
   End
   Begin VB.CheckBox Check19 
      Caption         =   "Check4"
      DataField       =   "b4"
      DataSource      =   "Adodc1"
      Height          =   195
      Left            =   5880
      TabIndex        =   47
      Top             =   5040
      Width           =   255
   End
   Begin VB.CheckBox Check18 
      Caption         =   "Check3"
      DataField       =   "b3"
      DataSource      =   "Adodc1"
      Height          =   195
      Left            =   4680
      TabIndex        =   46
      Top             =   5040
      Width           =   255
   End
   Begin VB.CheckBox Check17 
      Caption         =   "Check2"
      DataField       =   "b2"
      DataSource      =   "Adodc1"
      Height          =   195
      Left            =   3480
      TabIndex        =   45
      Top             =   5040
      Width           =   255
   End
   Begin VB.CheckBox Check16 
      Caption         =   "Check1"
      DataField       =   "b1"
      DataSource      =   "Adodc1"
      Height          =   195
      Left            =   2400
      TabIndex        =   44
      Top             =   5040
      Width           =   255
   End
   Begin VB.CheckBox Check15 
      Caption         =   "Check5"
      DataField       =   "c5"
      DataSource      =   "Adodc1"
      Height          =   195
      Left            =   6960
      TabIndex        =   43
      Top             =   5760
      Width           =   255
   End
   Begin VB.CheckBox Check14 
      Caption         =   "Check4"
      DataField       =   "c4"
      DataSource      =   "Adodc1"
      Height          =   195
      Left            =   5880
      TabIndex        =   42
      Top             =   5760
      Width           =   255
   End
   Begin VB.CheckBox Check13 
      Caption         =   "Check3"
      DataField       =   "c3"
      DataSource      =   "Adodc1"
      Height          =   195
      Left            =   4680
      TabIndex        =   41
      Top             =   5760
      Width           =   255
   End
   Begin VB.CheckBox Check12 
      Caption         =   "Check2"
      DataField       =   "c2"
      DataSource      =   "Adodc1"
      Height          =   195
      Left            =   3480
      TabIndex        =   40
      Top             =   5760
      Width           =   255
   End
   Begin VB.CheckBox Check11 
      Caption         =   "Check1"
      DataField       =   "c1"
      DataSource      =   "Adodc1"
      Height          =   195
      Left            =   2400
      TabIndex        =   39
      Top             =   5760
      Width           =   255
   End
   Begin VB.CheckBox Check10 
      Caption         =   "Check5"
      DataField       =   "d5"
      DataSource      =   "Adodc1"
      Height          =   195
      Left            =   6960
      TabIndex        =   38
      Top             =   6600
      Width           =   255
   End
   Begin VB.CheckBox Check9 
      Caption         =   "Check4"
      DataField       =   "d4"
      DataSource      =   "Adodc1"
      Height          =   195
      Left            =   5880
      TabIndex        =   37
      Top             =   6600
      Width           =   255
   End
   Begin VB.CheckBox Check8 
      Caption         =   "Check3"
      DataField       =   "d3"
      DataSource      =   "Adodc1"
      Height          =   195
      Left            =   4680
      TabIndex        =   36
      Top             =   6600
      Width           =   255
   End
   Begin VB.CheckBox Check7 
      Caption         =   "Check2"
      DataField       =   "d2"
      DataSource      =   "Adodc1"
      Height          =   195
      Left            =   3480
      TabIndex        =   35
      Top             =   6600
      Width           =   255
   End
   Begin VB.CheckBox Check6 
      Caption         =   "Check1"
      DataField       =   "d1"
      DataSource      =   "Adodc1"
      Height          =   195
      Left            =   2400
      TabIndex        =   34
      Top             =   6600
      Width           =   255
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   570
      Left            =   9000
      Top             =   840
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1005
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
      Connect         =   $"Form1.frx":0015
      OLEDBString     =   $"Form1.frx":00A2
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "servicecard"
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
   Begin VB.Frame Frame3 
      Caption         =   "miscellanous(optional)"
      Height          =   2175
      Left            =   240
      TabIndex        =   27
      Top             =   8880
      Width           =   8295
      Begin VB.TextBox Text7 
         DataField       =   "any other comment"
         DataSource      =   "Adodc1"
         Height          =   1095
         Left            =   2640
         TabIndex        =   31
         Top             =   960
         Width           =   5535
      End
      Begin VB.TextBox Text6 
         DataField       =   "how can we improve"
         DataSource      =   "Adodc1"
         Height          =   855
         Left            =   2640
         TabIndex        =   29
         Top             =   120
         Width           =   5535
      End
      Begin VB.Label Label21 
         Caption         =   "any comment etc you'd like to leave"
         Height          =   495
         Left            =   240
         TabIndex        =   30
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label Label20 
         Caption         =   "how can we improve ourselves"
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   3015
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "your experience here"
      Height          =   5175
      Left            =   240
      TabIndex        =   15
      Top             =   3600
      Width           =   8295
      Begin VB.Label Label19 
         Caption         =   "average"
         Height          =   375
         Left            =   4200
         TabIndex        =   26
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "food"
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "overall experience"
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   4320
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "cleanlines"
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "service"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "our rates for our food"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "environment"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label15 
         Caption         =   "excellent"
         Height          =   375
         Left            =   6600
         TabIndex        =   19
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label16 
         Caption         =   "good"
         Height          =   375
         Left            =   5520
         TabIndex        =   18
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label17 
         Caption         =   "poor"
         Height          =   375
         Left            =   3240
         TabIndex        =   17
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label18 
         Caption         =   "rubbish"
         Height          =   375
         Left            =   2040
         TabIndex        =   16
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "your personal details"
      Height          =   3135
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   8175
      Begin VB.ComboBox Combo1 
         DataField       =   "hw often visit"
         DataSource      =   "Adodc1"
         Height          =   315
         ItemData        =   "Form1.frx":012F
         Left            =   5880
         List            =   "Form1.frx":0142
         TabIndex        =   32
         Top             =   1800
         Width           =   2055
      End
      Begin VB.TextBox Text5 
         DataField       =   "mobile"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   5880
         TabIndex        =   14
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox Text4 
         DataField       =   "phone"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   5880
         TabIndex        =   13
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox Text3 
         DataField       =   "address"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   1560
         TabIndex        =   10
         Top             =   2520
         Width           =   6375
      End
      Begin VB.TextBox Text2 
         DataField       =   "age"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   1800
         TabIndex        =   3
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         DataField       =   "name"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   1800
         TabIndex        =   2
         Top             =   480
         Width           =   2175
      End
      Begin VB.ComboBox Combo5 
         DataField       =   "gender"
         DataSource      =   "Adodc1"
         Height          =   315
         ItemData        =   "Form1.frx":017E
         Left            =   1800
         List            =   "Form1.frx":0188
         TabIndex        =   1
         ToolTipText     =   $"Form1.frx":019A
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label Label22 
         Caption         =   "how often do you visit"
         Height          =   375
         Left            =   4080
         TabIndex        =   33
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   "mobile number"
         Height          =   495
         Left            =   4200
         TabIndex        =   12
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "phone number"
         Height          =   375
         Left            =   4200
         TabIndex        =   11
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "address"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "age"
         Height          =   375
         Left            =   480
         TabIndex        =   8
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label4 
         Height          =   375
         Left            =   4080
         TabIndex        =   7
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "gender"
         Height          =   375
         Left            =   480
         TabIndex        =   6
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "your name"
         Height          =   375
         Left            =   480
         TabIndex        =   5
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "nationality"
         Height          =   375
         Left            =   4080
         TabIndex        =   4
         Top             =   600
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Option1_Click()

End Sub

Private Sub Option12_Click()

End Sub

Private Sub Option17_Click()

End Sub

Private Sub Option27_Click()

End Sub

Private Sub Command1_Click()
On Error GoTo err
Adodc1.Recordset.Save
Exit Sub
err:
End Sub

Private Sub Command2_Click()
On Error GoTo err
Adodc1.Recordset.Delete
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
Exit Sub
err:
End Sub

Private Sub Command5_Click()

Check6.Value = Unchecked
Check7.Value = Unchecked
Check8.Value = Unchecked
Check9.Value = Unchecked
Check10.Value = Unchecked
Check11.Value = Unchecked
Check12.Value = Unchecked
Check13.Value = Unchecked
Check14.Value = Unchecked
Check15.Value = Unchecked
Check16.Value = Unchecked
Check17.Value = Unchecked
Check18.Value = Unchecked
Check19.Value = Unchecked
Check20.Value = Unchecked
Check21.Value = Unchecked
Check22.Value = Unchecked
Check23.Value = Unchecked
Check23.Value = Unchecked
Check24.Value = Unchecked
Check25.Value = Unchecked
Check26.Value = Unchecked
Check27.Value = Unchecked
Check28.Value = Unchecked
Check29.Value = Unchecked
Check30.Value = Unchecked
Check31.Value = Unchecked
Check32.Value = Unchecked
Check33.Value = Unchecked
Check34.Value = Unchecked
Check35.Value = Unchecked
End Sub

Private Sub Command6_Click()
If Text8.Text = "" Then
MsgBox "enter value for searching a record"

Else

    str1 = Text8.Text
    strsearch = "name like '" & str1 & "'"

    Adodc1.Recordset.MoveFirst
    Adodc1.Recordset.Find (strsearch)
    MsgBox "results found"
End If
End Sub
