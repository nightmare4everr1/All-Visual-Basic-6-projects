VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9885
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13815
   LinkTopic       =   "Form1"
   ScaleHeight     =   9885
   ScaleWidth      =   13815
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text12 
      Height          =   615
      Left            =   8400
      TabIndex        =   28
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "search"
      Height          =   615
      Left            =   6720
      TabIndex        =   27
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "refresh"
      Height          =   495
      Left            =   5160
      TabIndex        =   26
      Top             =   3240
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   5160
      Top             =   720
      Width           =   1695
      _ExtentX        =   2990
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Microsoft Visual Studio\VB98\football\football.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Microsoft Visual Studio\VB98\football\football.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "badminton"
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
   Begin VB.CommandButton Command3 
      Caption         =   "delete"
      Height          =   495
      Left            =   5160
      TabIndex        =   25
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "save"
      Height          =   495
      Left            =   5160
      TabIndex        =   24
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "add new"
      Height          =   495
      Left            =   5160
      TabIndex        =   23
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      DataField       =   "name"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3240
      TabIndex        =   22
      Top             =   840
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form1.frx":0000
      Height          =   2415
      Left            =   0
      TabIndex        =   20
      Top             =   7440
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   4260
      _Version        =   393216
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
      ColumnCount     =   11
      BeginProperty Column00 
         DataField       =   "speed"
         Caption         =   "speed"
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
         DataField       =   "attacking"
         Caption         =   "attacking"
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
         DataField       =   "defending"
         Caption         =   "defending"
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
         DataField       =   "dodging"
         Caption         =   "dodging"
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
         DataField       =   "aim"
         Caption         =   "aim"
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
         DataField       =   "loner or team"
         Caption         =   "loner or team"
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
         DataField       =   "snatching skills"
         Caption         =   "snatching skills"
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
         DataField       =   "lob"
         Caption         =   "lob"
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
         DataField       =   "shoot"
         Caption         =   "shoot"
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
         DataField       =   "passing"
         Caption         =   "passing"
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
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text11 
      DataField       =   "attacking"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3240
      TabIndex        =   19
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox Text10 
      DataField       =   "defending"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3240
      TabIndex        =   18
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox Text9 
      DataField       =   "dodging"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3240
      TabIndex        =   17
      Top             =   3240
      Width           =   1575
   End
   Begin VB.TextBox Text8 
      DataField       =   "aim"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3240
      TabIndex        =   16
      Top             =   3840
      Width           =   1575
   End
   Begin VB.TextBox Text7 
      DataField       =   "loner or team"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3240
      TabIndex        =   15
      Top             =   4440
      Width           =   1575
   End
   Begin VB.TextBox Text6 
      DataField       =   "snatching skills"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3240
      TabIndex        =   14
      Top             =   5040
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      DataField       =   "lob"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3240
      TabIndex        =   13
      Top             =   5640
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      DataField       =   "shoot"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3240
      TabIndex        =   12
      Top             =   6240
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      DataField       =   "passing"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3240
      TabIndex        =   11
      Top             =   6840
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      DataField       =   "speed"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3240
      TabIndex        =   10
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label11 
      Caption         =   "name"
      Height          =   615
      Left            =   960
      TabIndex        =   21
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label10 
      Caption         =   "passing"
      Height          =   495
      Left            =   960
      TabIndex        =   9
      Top             =   6840
      Width           =   1815
   End
   Begin VB.Label Label9 
      Caption         =   "shoot"
      Height          =   495
      Left            =   960
      TabIndex        =   8
      Top             =   6240
      Width           =   2175
   End
   Begin VB.Label Label8 
      Caption         =   "lob"
      Height          =   495
      Left            =   960
      TabIndex        =   7
      Top             =   5760
      Width           =   2295
   End
   Begin VB.Label Label7 
      Caption         =   "attacking"
      Height          =   495
      Left            =   960
      TabIndex        =   6
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "defending"
      Height          =   495
      Left            =   960
      TabIndex        =   5
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "dodging"
      Height          =   495
      Left            =   960
      TabIndex        =   4
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "aim"
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "loner or team"
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "snatching skills"
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   5160
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "speed"
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   1560
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo err
Adodc1.Recordset.AddNew
Exit Sub
err:
End Sub

Private Sub Command2_Click()
On Error GoTo err
Adodc1.Recordset.Save
Exit Sub
err:
End Sub

Private Sub Command3_Click()
On Error GoTo err
Adodc1.Recordset.Delete
Exit Sub
err:
End Sub

Private Sub Command4_Click()
Unload Form1
Load Form1
Form1.Show
End Sub

Private Sub Command5_Click()
If Text12.Text = "" Then
MsgBox "enter value for searching a record"

Else

    str1 = Text12.Text
    strsearch = "name like '" & str1 & "'"

    Adodc1.Recordset.MoveFirst
    Adodc1.Recordset.Find (strsearch)
    MsgBox "search complete"
End If
End Sub
