VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13455
   LinkTopic       =   "Form1"
   ScaleHeight     =   9915
   ScaleWidth      =   13455
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   11055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15165
      _ExtentX        =   26749
      _ExtentY        =   19500
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "Form1.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Text12"
      Tab(0).Control(1)=   "Command5"
      Tab(0).Control(2)=   "Command4"
      Tab(0).Control(3)=   "Command3"
      Tab(0).Control(4)=   "Command2"
      Tab(0).Control(5)=   "Command1"
      Tab(0).Control(6)=   "Text2"
      Tab(0).Control(7)=   "Text11"
      Tab(0).Control(8)=   "Text10"
      Tab(0).Control(9)=   "Text9"
      Tab(0).Control(10)=   "Text8"
      Tab(0).Control(11)=   "Text7"
      Tab(0).Control(12)=   "Text6"
      Tab(0).Control(13)=   "Text5"
      Tab(0).Control(14)=   "Text4"
      Tab(0).Control(15)=   "Text3"
      Tab(0).Control(16)=   "Text1"
      Tab(0).Control(17)=   "Adodc1"
      Tab(0).Control(18)=   "DataGrid1"
      Tab(0).Control(19)=   "Label11"
      Tab(0).Control(20)=   "Label10"
      Tab(0).Control(21)=   "Label9"
      Tab(0).Control(22)=   "Label8"
      Tab(0).Control(23)=   "Label7"
      Tab(0).Control(24)=   "Label6"
      Tab(0).Control(25)=   "Label5"
      Tab(0).Control(26)=   "Label4"
      Tab(0).Control(27)=   "Label3"
      Tab(0).Control(28)=   "Label2"
      Tab(0).Control(29)=   "Label1"
      Tab(0).ControlCount=   30
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "Form1.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label12"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label13"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label14"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label15"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label16"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label17"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label18"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label19"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label20"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label21"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label22"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label23"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label24"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Adodc2"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "DataGrid2"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Text13"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Text14"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Text15"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Text16"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Text17"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Text18"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Text19"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Text20"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Text21"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "Text22"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "Text23"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "Command6"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "Command7"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "Command8"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "Command9"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "Text24"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "Timer1"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "Timer2"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "Command10"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "Text25"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).ControlCount=   35
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "Form1.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Command11"
      Tab(2).ControlCount=   1
      Begin VB.CommandButton Command11 
         Caption         =   "goto form2"
         Height          =   975
         Left            =   -74400
         TabIndex        =   62
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox Text25 
         Height          =   495
         Left            =   12360
         TabIndex        =   47
         Top             =   1200
         Width           =   1815
      End
      Begin VB.CommandButton Command10 
         Caption         =   "search"
         Height          =   615
         Left            =   10080
         TabIndex        =   46
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Timer Timer2 
         Interval        =   1
         Left            =   6720
         Top             =   480
      End
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   5040
         Top             =   5040
      End
      Begin VB.TextBox Text24 
         DataField       =   "percentage"
         DataSource      =   "Adodc2"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   2040
         TabIndex        =   45
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton Command9 
         Caption         =   "refresh"
         Height          =   495
         Left            =   6240
         TabIndex        =   44
         ToolTipText     =   "use it if encountering a problem"
         Top             =   3000
         Width           =   1335
      End
      Begin VB.CommandButton Command8 
         Caption         =   "delete"
         Height          =   495
         Left            =   4680
         TabIndex        =   43
         ToolTipText     =   "permanently deletes one form"
         Top             =   3000
         Width           =   1455
      End
      Begin VB.CommandButton Command7 
         Caption         =   "add new"
         Height          =   495
         Left            =   6240
         TabIndex        =   42
         ToolTipText     =   "creates a blank form"
         Top             =   2400
         Width           =   1335
      End
      Begin VB.CommandButton Command6 
         Caption         =   "save"
         Height          =   495
         Left            =   4680
         TabIndex        =   41
         ToolTipText     =   "saves the data"
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox Text23 
         DataField       =   "your age"
         DataSource      =   "Adodc2"
         Height          =   495
         Left            =   3840
         TabIndex        =   40
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox Text22 
         DataField       =   "specialise in"
         DataSource      =   "Adodc2"
         Height          =   495
         Left            =   2280
         TabIndex        =   39
         Top             =   5520
         Width           =   3855
      End
      Begin VB.TextBox Text21 
         DataField       =   "aggresive, defense, medium"
         DataSource      =   "Adodc2"
         Height          =   495
         Left            =   2280
         TabIndex        =   38
         Top             =   5040
         Width           =   1455
      End
      Begin VB.TextBox Text20 
         DataField       =   "reflexes"
         DataSource      =   "Adodc2"
         Height          =   525
         Left            =   2280
         TabIndex        =   37
         Top             =   4560
         Width           =   1455
      End
      Begin VB.TextBox Text19 
         DataField       =   "serve"
         DataSource      =   "Adodc2"
         Height          =   495
         Left            =   2280
         TabIndex        =   36
         Top             =   4080
         Width           =   1455
      End
      Begin VB.TextBox Text18 
         DataField       =   "stamina"
         DataSource      =   "Adodc2"
         Height          =   495
         Left            =   2280
         TabIndex        =   35
         Top             =   3600
         Width           =   1455
      End
      Begin VB.TextBox Text17 
         DataField       =   "what is ur avg rally"
         DataSource      =   "Adodc2"
         Height          =   495
         Left            =   2280
         TabIndex        =   34
         Top             =   3120
         Width           =   1455
      End
      Begin VB.TextBox Text16 
         DataField       =   "net shots"
         DataSource      =   "Adodc2"
         Height          =   495
         Left            =   2280
         TabIndex        =   33
         Top             =   2640
         Width           =   1455
      End
      Begin VB.TextBox Text15 
         DataField       =   "backhand"
         DataSource      =   "Adodc2"
         Height          =   495
         Left            =   2280
         TabIndex        =   32
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox Text14 
         DataField       =   "forehand"
         DataSource      =   "Adodc2"
         Height          =   495
         Left            =   2280
         TabIndex        =   31
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox Text13 
         DataField       =   "name"
         DataSource      =   "Adodc2"
         Height          =   495
         Left            =   2280
         TabIndex        =   30
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox Text12 
         Height          =   615
         Left            =   -66600
         TabIndex        =   17
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CommandButton Command5 
         Caption         =   "search"
         Height          =   615
         Left            =   -68280
         TabIndex        =   16
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CommandButton Command4 
         Caption         =   "refresh"
         Height          =   495
         Left            =   -69840
         TabIndex        =   15
         Top             =   3360
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "delete"
         Height          =   495
         Left            =   -69840
         TabIndex        =   14
         Top             =   2760
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "save"
         Height          =   495
         Left            =   -69840
         TabIndex        =   13
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "add new"
         Height          =   495
         Left            =   -69840
         TabIndex        =   12
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         DataField       =   "name"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   -71760
         TabIndex        =   11
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox Text11 
         DataField       =   "attacking"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   -71760
         TabIndex        =   10
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox Text10 
         DataField       =   "defending"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   -71760
         TabIndex        =   9
         Top             =   2760
         Width           =   1575
      End
      Begin VB.TextBox Text9 
         DataField       =   "dodging"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   -71760
         TabIndex        =   8
         Top             =   3360
         Width           =   1575
      End
      Begin VB.TextBox Text8 
         DataField       =   "aim"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   -71760
         TabIndex        =   7
         Top             =   3960
         Width           =   1575
      End
      Begin VB.TextBox Text7 
         DataField       =   "loner or team"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   -71760
         TabIndex        =   6
         Top             =   4560
         Width           =   1575
      End
      Begin VB.TextBox Text6 
         DataField       =   "snatching skills"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   -71760
         TabIndex        =   5
         Top             =   5160
         Width           =   1575
      End
      Begin VB.TextBox Text5 
         DataField       =   "lob"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   -71760
         TabIndex        =   4
         Top             =   5760
         Width           =   1575
      End
      Begin VB.TextBox Text4 
         DataField       =   "shoot"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   -71760
         TabIndex        =   3
         Top             =   6360
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         DataField       =   "passing"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   -71760
         TabIndex        =   2
         Top             =   6960
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         DataField       =   "speed"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   -71760
         TabIndex        =   1
         Top             =   1560
         Width           =   1575
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   495
         Left            =   -69840
         Top             =   840
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
         Connect         =   $"Form1.frx":0054
         OLEDBString     =   $"Form1.frx":00E5
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "football"
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
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Form1.frx":0176
         Height          =   2415
         Left            =   -75000
         TabIndex        =   18
         Top             =   7560
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
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "Form1.frx":018B
         Height          =   4095
         Left            =   0
         TabIndex        =   48
         Top             =   6120
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
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   495
         Left            =   4680
         Top             =   3720
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
         Connect         =   $"Form1.frx":01A0
         OLEDBString     =   $"Form1.frx":0233
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
      Begin VB.Label Label24 
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
         Left            =   360
         TabIndex        =   61
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label23 
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
         Left            =   360
         TabIndex        =   60
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label22 
         Caption         =   "forehand"
         Height          =   375
         Left            =   1080
         TabIndex        =   59
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label21 
         Caption         =   "backhand"
         Height          =   375
         Left            =   1080
         TabIndex        =   58
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label20 
         Caption         =   "net shots"
         Height          =   375
         Left            =   1080
         TabIndex        =   57
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label19 
         Caption         =   "what is your avg rally"
         Height          =   375
         Left            =   1080
         TabIndex        =   56
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label18 
         Caption         =   "stamina"
         Height          =   375
         Left            =   1080
         TabIndex        =   55
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label Label17 
         Caption         =   "serve"
         Height          =   375
         Left            =   1080
         TabIndex        =   54
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Label Label16 
         Caption         =   "reflexes"
         Height          =   375
         Left            =   1080
         TabIndex        =   53
         Top             =   4560
         Width           =   1215
      End
      Begin VB.Label Label15 
         Caption         =   "aggresive,defensive,medium"
         Height          =   375
         Left            =   0
         TabIndex        =   52
         Top             =   5040
         Width           =   2175
      End
      Begin VB.Label Label14 
         Caption         =   "specialise in"
         Height          =   375
         Left            =   1080
         TabIndex        =   51
         Top             =   5520
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "age"
         Height          =   375
         Left            =   3960
         TabIndex        =   50
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "name"
         Height          =   375
         Left            =   1080
         TabIndex        =   49
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "name"
         Height          =   615
         Left            =   -74040
         TabIndex        =   29
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label10 
         Caption         =   "passing"
         Height          =   495
         Left            =   -74040
         TabIndex        =   28
         Top             =   6960
         Width           =   1815
      End
      Begin VB.Label Label9 
         Caption         =   "shoot"
         Height          =   495
         Left            =   -74040
         TabIndex        =   27
         Top             =   6360
         Width           =   2175
      End
      Begin VB.Label Label8 
         Caption         =   "lob"
         Height          =   495
         Left            =   -74040
         TabIndex        =   26
         Top             =   5880
         Width           =   2295
      End
      Begin VB.Label Label7 
         Caption         =   "attacking"
         Height          =   495
         Left            =   -74040
         TabIndex        =   25
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label Label6 
         Caption         =   "defending"
         Height          =   495
         Left            =   -74040
         TabIndex        =   24
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "dodging"
         Height          =   495
         Left            =   -74040
         TabIndex        =   23
         Top             =   3480
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "aim"
         Height          =   495
         Left            =   -74040
         TabIndex        =   22
         Top             =   4080
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "loner or team"
         Height          =   495
         Left            =   -74040
         TabIndex        =   21
         Top             =   4680
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "snatching skills"
         Height          =   495
         Left            =   -74040
         TabIndex        =   20
         Top             =   5280
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "speed"
         Height          =   495
         Left            =   -74040
         TabIndex        =   19
         Top             =   1680
         Width           =   1935
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command11_Click()
Form1.Hide
Form2.Show
End Sub

Private Sub Command6_Click()
On Error GoTo err
Adodc1.Recordset.Save
Exit Sub
err:
Unload Form1
Load Form1
Form1.Show
End Sub

Private Sub Command7_Click()
On Error GoTo err
Adodc1.Recordset.AddNew
Exit Sub
err:
End Sub

Private Sub Command8_Click()
On Error GoTo err
Adodc1.Recordset.Delete
Exit Sub
err:
Unload Form1
Load Form1
Form1.Show
End Sub

Private Sub Command9_Click()
Unload Form1
Load Form1
Form1.Show
End Sub

Private Sub Command10_Click()
If Text25.Text = "" Then
MsgBox "enter value for searching a record"

Else

    str1 = Text25.Text
    strsearch = "name like '" & str1 & "'"

    Adodc1.Recordset.MoveFirst
    Adodc1.Recordset.Find (strsearch)
    MsgBox "search complete"
End If
End Sub

Private Sub Timer1_Timer()
Text12.Text = (Val(Text2.Text) + Val(Text3.Text) + Val(Text4.Text) + Val(Text6.Text) + Val(Text7.Text) + Val(Text8.Text)) * 100 / 72
End Sub


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

