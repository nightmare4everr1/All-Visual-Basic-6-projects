VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form1.frx":0000
      Height          =   1815
      Left            =   1200
      TabIndex        =   8
      Top             =   6600
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   3201
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
   Begin VB.TextBox Text4 
      Height          =   1215
      Left            =   3240
      TabIndex        =   7
      Top             =   4920
      Width           =   2655
   End
   Begin VB.CommandButton Command4 
      Caption         =   "search"
      Height          =   1215
      Left            =   600
      TabIndex        =   6
      Top             =   4920
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "delete"
      Height          =   975
      Left            =   4440
      TabIndex        =   5
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "addnew"
      Height          =   975
      Left            =   2400
      TabIndex        =   4
      Top             =   3000
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "save"
      Height          =   975
      Left            =   240
      TabIndex        =   3
      Top             =   3000
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      DataField       =   "pocketmoney"
      DataSource      =   "Adodc1"
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   2160
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      DataField       =   "age"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1200
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      DataField       =   "name"
      DataSource      =   "Adodc1"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   2775
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   960
      Top             =   4080
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   1296
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Documents and Settings\ADNAN\My Documents\student.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Documents and Settings\ADNAN\My Documents\student.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "student"
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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Data1_Validate(Action As Integer, Save As Integer)

End Sub

Private Sub Command1_Click()
Adodc1.Recordset.Save
Adodc1.Refresh
End Sub

Private Sub Command2_Click()
Adodc1.Recordset.AddNew
End Sub

Private Sub Command3_Click()
Adodc1.Recordset.Delete
Adodc1.Update
End Sub

Private Sub OLE1_Updated(Code As Integer)

End Sub

Private Sub Command4_Click()
Form1.Hide
Form2.Show
    
End Sub

