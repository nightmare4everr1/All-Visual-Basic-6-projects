VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Ahmed Adnan's Rocket Velocity Project"
   ClientHeight    =   11010
   ClientLeft      =   540
   ClientTop       =   690
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer11 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   240
      Top             =   9000
   End
   Begin VB.Frame Frame8 
      Height          =   11055
      Left            =   -120
      TabIndex        =   115
      Top             =   10800
      Width           =   15255
      Begin VB.CommandButton Command16 
         Caption         =   "Previous Setting"
         Height          =   735
         Left            =   4680
         TabIndex        =   120
         Top             =   3720
         Width           =   2295
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Next Setting"
         Height          =   735
         Left            =   8280
         TabIndex        =   119
         Top             =   3720
         Width           =   2535
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Proceed To Project"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   6000
         TabIndex        =   116
         Top             =   4800
         Width           =   4095
      End
      Begin VB.Label Label52 
         Caption         =   $"Form1.frx":0000
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4095
         Left            =   600
         TabIndex        =   121
         Top             =   6600
         Width           =   14415
      End
      Begin VB.Label Label51 
         Caption         =   "Choose Your Settings"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   48
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   2040
         TabIndex        =   118
         Top             =   240
         Width           =   9855
      End
      Begin VB.Label Label50 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         DataField       =   "name"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1560
         TabIndex        =   117
         Top             =   2160
         Width           =   13815
      End
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Stop"
      Height          =   375
      Left            =   1080
      TabIndex        =   102
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Start"
      Height          =   375
      Left            =   0
      TabIndex        =   101
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text18 
      Height          =   375
      Left            =   360
      TabIndex        =   100
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4815
      Left            =   4440
      TabIndex        =   53
      Top             =   1800
      Visible         =   0   'False
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   8493
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   -2147483637
      ForeColor       =   49152
      TabCaption(0)   =   "settings"
      TabPicture(0)   =   "Form1.frx":00BD
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "values"
      TabPicture(1)   =   "Form1.frx":00D9
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Adodc1"
      Tab(1).Control(1)=   "Frame6"
      Tab(1).Control(2)=   "Frame7"
      Tab(1).Control(3)=   "Timer8"
      Tab(1).Control(4)=   "Timer9"
      Tab(1).Control(5)=   "Command11"
      Tab(1).Control(6)=   "Command12"
      Tab(1).Control(7)=   "Command13"
      Tab(1).Control(8)=   "Text21"
      Tab(1).ControlCount=   9
      Begin VB.TextBox Text21 
         DataField       =   "name"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   -71760
         TabIndex        =   113
         Text            =   "Text21"
         Top             =   5640
         Width           =   4215
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Delete"
         Height          =   375
         Left            =   -69000
         TabIndex        =   112
         ToolTipText     =   "permanently deletes the record"
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Save"
         Height          =   375
         Left            =   -70560
         TabIndex        =   111
         ToolTipText     =   "saves the values you have inputted.it will overwrite on existing data if you havent added a new record"
         Top             =   1080
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Add"
         Height          =   375
         Left            =   -70560
         TabIndex        =   110
         ToolTipText     =   "Adds a blank new record so you can store your values"
         Top             =   600
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Timer Timer9 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   -70920
         Top             =   240
      End
      Begin VB.Timer Timer8 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   -70200
         Top             =   120
      End
      Begin VB.Frame Frame7 
         Caption         =   "User Defined Settings"
         Height          =   5295
         Left            =   -74880
         TabIndex        =   65
         Top             =   1680
         Visible         =   0   'False
         Width           =   7935
         Begin VB.TextBox Text17 
            DataField       =   "P"
            DataSource      =   "Adodc1"
            Height          =   285
            Left            =   3240
            TabIndex        =   93
            Top             =   840
            Width           =   975
         End
         Begin VB.TextBox Text16 
            DataField       =   "I"
            DataSource      =   "Adodc1"
            Height          =   285
            Left            =   3240
            TabIndex        =   92
            Top             =   1080
            Width           =   975
         End
         Begin VB.TextBox Text15 
            DataField       =   "S"
            DataSource      =   "Adodc1"
            Height          =   285
            Left            =   3240
            TabIndex        =   91
            Top             =   1320
            Width           =   975
         End
         Begin VB.TextBox Text14 
            DataField       =   "Cx"
            DataSource      =   "Adodc1"
            Height          =   285
            Left            =   3240
            TabIndex        =   90
            Top             =   1560
            Width           =   975
         End
         Begin VB.TextBox Text13 
            DataField       =   "T"
            DataSource      =   "Adodc1"
            Height          =   495
            Left            =   6000
            TabIndex        =   83
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox Text12 
            DataField       =   "G"
            DataSource      =   "Adodc1"
            Height          =   495
            Left            =   6000
            TabIndex        =   82
            Top             =   1440
            Width           =   1215
         End
         Begin VB.TextBox Text11 
            DataField       =   "m"
            DataSource      =   "Adodc1"
            Height          =   495
            Left            =   6000
            TabIndex        =   81
            Top             =   2040
            Width           =   1215
         End
         Begin VB.TextBox Text10 
            DataField       =   "stepb"
            DataSource      =   "Adodc1"
            Height          =   495
            Left            =   6000
            TabIndex        =   80
            Top             =   2640
            Width           =   1215
         End
         Begin VB.TextBox Text9 
            DataField       =   "vo"
            DataSource      =   "Adodc1"
            Height          =   495
            Left            =   6000
            TabIndex        =   79
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox Text8 
            DataField       =   "stepa"
            DataSource      =   "Adodc1"
            Height          =   495
            Left            =   6000
            TabIndex        =   78
            Top             =   3240
            Width           =   1215
         End
         Begin VB.TextBox Text7 
            DataField       =   "mt"
            DataSource      =   "Adodc1"
            Height          =   495
            Left            =   1560
            TabIndex        =   71
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox Text6 
            DataField       =   "mo"
            DataSource      =   "Adodc1"
            Height          =   495
            Left            =   1560
            TabIndex        =   70
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox Text5 
            DataField       =   "mp"
            DataSource      =   "Adodc1"
            Height          =   495
            Left            =   1560
            TabIndex        =   69
            Top             =   1440
            Width           =   1215
         End
         Begin VB.TextBox Text4 
            DataField       =   "u"
            DataSource      =   "Adodc1"
            Height          =   495
            Left            =   1560
            TabIndex        =   68
            Top             =   2040
            Width           =   1215
         End
         Begin VB.TextBox Text3 
            DataField       =   "ux"
            DataSource      =   "Adodc1"
            Height          =   495
            Left            =   1560
            TabIndex        =   67
            Top             =   2640
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            DataField       =   "vm"
            DataSource      =   "Adodc1"
            Height          =   495
            Left            =   1560
            TabIndex        =   66
            Top             =   3240
            Width           =   1215
         End
         Begin VB.Label Label49 
            Caption         =   "Name Of Data"
            Height          =   375
            Left            =   1560
            TabIndex        =   114
            Top             =   3960
            Width           =   2055
         End
         Begin VB.Label Label43 
            BackStyle       =   0  'Transparent
            Caption         =   "P"
            Height          =   255
            Left            =   2880
            TabIndex        =   97
            Top             =   840
            Width           =   495
         End
         Begin VB.Label Label42 
            BackStyle       =   0  'Transparent
            Caption         =   "I"
            Height          =   255
            Left            =   2880
            TabIndex        =   96
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label41 
            BackStyle       =   0  'Transparent
            Caption         =   "S"
            Height          =   255
            Left            =   2880
            TabIndex        =   95
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label40 
            BackStyle       =   0  'Transparent
            Caption         =   "Cx"
            Height          =   255
            Left            =   2880
            TabIndex        =   94
            Top             =   1560
            Width           =   735
         End
         Begin VB.Label Label39 
            BackStyle       =   0  'Transparent
            Caption         =   "T"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4440
            TabIndex        =   89
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label38 
            BackStyle       =   0  'Transparent
            Caption         =   "gravity"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4440
            TabIndex        =   88
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label Label37 
            BackStyle       =   0  'Transparent
            Caption         =   "m"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4440
            TabIndex        =   87
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label Label36 
            BackStyle       =   0  'Transparent
            Caption         =   "step3"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4440
            TabIndex        =   86
            Top             =   2640
            Width           =   1455
         End
         Begin VB.Label Label35 
            BackStyle       =   0  'Transparent
            Caption         =   "v(0)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4440
            TabIndex        =   85
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label34 
            BackStyle       =   0  'Transparent
            Caption         =   "step2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4440
            TabIndex        =   84
            Top             =   3240
            Width           =   1335
         End
         Begin VB.Label Label26 
            BackStyle       =   0  'Transparent
            Caption         =   "m(t)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   77
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label25 
            BackStyle       =   0  'Transparent
            Caption         =   "u^x"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   76
            Top             =   2640
            Width           =   1215
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "v(m)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   75
            Top             =   3240
            Width           =   1215
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "u"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   74
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "m(0)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   73
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "m(p)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   72
            Top             =   1440
            Width           =   1215
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Enter password to proceed"
         Height          =   1095
         Left            =   -74760
         TabIndex        =   62
         Top             =   480
         Width           =   4095
         Begin VB.CommandButton Command9 
            Caption         =   "close"
            Height          =   375
            Left            =   2160
            TabIndex        =   98
            Top             =   240
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.CommandButton Command8 
            Caption         =   "password"
            Height          =   375
            Left            =   120
            TabIndex        =   64
            Top             =   240
            Width           =   1935
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   120
            TabIndex        =   63
            Top             =   720
            Width           =   1935
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Settings"
         Height          =   4095
         Left            =   120
         TabIndex        =   54
         Top             =   360
         Width           =   5055
         Begin VB.CommandButton Command17 
            Caption         =   "Help"
            Height          =   375
            Left            =   3600
            TabIndex        =   122
            Top             =   3600
            Width           =   1335
         End
         Begin VB.CommandButton Command10 
            Caption         =   "change background music"
            Height          =   495
            Left            =   1200
            TabIndex        =   99
            Top             =   2880
            Width           =   2295
         End
         Begin VB.Frame Frame1 
            Height          =   1935
            Left            =   480
            TabIndex        =   58
            Top             =   240
            Width           =   3975
            Begin VB.OptionButton Option1 
               Caption         =   "Manually"
               Height          =   255
               Left            =   360
               TabIndex        =   60
               Top             =   1200
               Value           =   -1  'True
               Width           =   1095
            End
            Begin VB.OptionButton Option2 
               Caption         =   "Automatically"
               Height          =   255
               Left            =   2040
               TabIndex        =   59
               Top             =   1200
               Width           =   1275
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "Do you want to imput v(m) and m(p) manually or do you want the computer to evalute them on pre-existing data"
               Height          =   615
               Left            =   240
               TabIndex        =   61
               Top             =   360
               Width           =   3015
            End
            Begin VB.Shape Shape1 
               BorderColor     =   &H8000000D&
               BorderWidth     =   5
               Height          =   495
               Left            =   240
               Top             =   1080
               Width           =   1335
            End
            Begin VB.Shape Shape2 
               BorderColor     =   &H8000000D&
               BorderWidth     =   5
               Height          =   495
               Left            =   1920
               Top             =   1080
               Visible         =   0   'False
               Width           =   1455
            End
         End
         Begin VB.Frame Frame5 
            Height          =   1215
            Left            =   480
            TabIndex        =   56
            Top             =   2280
            Width           =   3975
            Begin VB.ComboBox Combo1 
               Height          =   315
               ItemData        =   "Form1.frx":00F5
               Left            =   240
               List            =   "Form1.frx":00FF
               TabIndex        =   57
               Text            =   "Click Me"
               Top             =   240
               Width           =   3015
            End
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Submit Changes"
            Height          =   375
            Left            =   240
            TabIndex        =   55
            Top             =   3600
            Width           =   3135
         End
         Begin VB.Timer Timer1 
            Interval        =   1
            Left            =   3720
            Top             =   2520
         End
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   -69000
         ToolTipText     =   "Scroll Through your previous settings"
         Top             =   1080
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
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
         Connect         =   $"Form1.frx":013D
         OLEDBString     =   $"Form1.frx":01C6
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "calendar"
         Caption         =   "Scroll"
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
   Begin VB.Timer Timer6 
      Interval        =   1
      Left            =   4560
      Top             =   8280
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4560
      Top             =   7800
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4080
      Top             =   8280
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4080
      Top             =   7800
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   240
      Top             =   2880
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Settings"
      Height          =   495
      Left            =   480
      TabIndex        =   45
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Height          =   3975
      Left            =   360
      TabIndex        =   37
      Top             =   2400
      Width           =   3375
      Begin VB.TextBox Text20 
         Height          =   285
         Left            =   1680
         TabIndex        =   107
         Top             =   3120
         Width           =   1335
      End
      Begin VB.TextBox Text19 
         Height          =   285
         Left            =   1440
         TabIndex        =   104
         Text            =   "1"
         Top             =   1080
         Width           =   375
      End
      Begin VB.Timer Timer10 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   2040
         Top             =   3360
      End
      Begin VB.Timer Timer7 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   2640
         Top             =   3360
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Draw V-T"
         Height          =   615
         Left            =   2640
         TabIndex        =   52
         Top             =   840
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Frame Frame3 
         Height          =   1575
         Left            =   360
         TabIndex        =   39
         Top             =   1440
         Width           =   2655
         Begin VB.ListBox List1 
            Height          =   1035
            ItemData        =   "Form1.frx":024F
            Left            =   120
            List            =   "Form1.frx":0251
            TabIndex        =   40
            Top             =   360
            Width           =   2415
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "velocity"
            Height          =   495
            Left            =   600
            TabIndex        =   42
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "time"
            Height          =   495
            Left            =   120
            TabIndex        =   41
            Top             =   120
            Width           =   1215
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "calculate velocity"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   38
         ToolTipText     =   "Calculate velocity-time profile"
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label47 
         Caption         =   "Label47"
         Height          =   255
         Left            =   120
         TabIndex        =   109
         Top             =   3480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label48 
         Caption         =   "Label48"
         Height          =   255
         Left            =   960
         TabIndex        =   108
         Top             =   3480
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label46 
         Caption         =   "Average Value"
         Height          =   255
         Left            =   240
         TabIndex        =   106
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Label Label45 
         Caption         =   "Second(s)"
         Height          =   255
         Left            =   1920
         TabIndex        =   105
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label44 
         Caption         =   "With intervals of"
         Height          =   495
         Left            =   240
         TabIndex        =   103
         Top             =   1080
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "reset"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   12240
      TabIndex        =   33
      Top             =   3720
      Width           =   1455
   End
   Begin VB.TextBox step2 
      Height          =   495
      Left            =   10800
      TabIndex        =   31
      Top             =   8520
      Width           =   1215
   End
   Begin VB.TextBox C 
      Height          =   285
      Left            =   12600
      TabIndex        =   27
      Top             =   5160
      Width           =   975
   End
   Begin VB.TextBox S 
      Height          =   285
      Left            =   12600
      TabIndex        =   26
      Top             =   4920
      Width           =   975
   End
   Begin VB.TextBox v0 
      Height          =   495
      Left            =   10800
      TabIndex        =   25
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox step3 
      Height          =   495
      Left            =   10800
      TabIndex        =   23
      Top             =   7920
      Width           =   1215
   End
   Begin VB.TextBox m 
      Height          =   495
      Left            =   10800
      TabIndex        =   21
      Top             =   7320
      Width           =   1215
   End
   Begin VB.TextBox G 
      Height          =   495
      Left            =   10800
      TabIndex        =   19
      Top             =   6720
      Width           =   1215
   End
   Begin VB.TextBox I 
      Height          =   285
      Left            =   12480
      TabIndex        =   15
      Top             =   3360
      Width           =   855
   End
   Begin VB.TextBox P 
      Height          =   285
      Left            =   12480
      TabIndex        =   14
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox T 
      Height          =   495
      Left            =   10800
      TabIndex        =   13
      Top             =   6120
      Width           =   1215
   End
   Begin VB.TextBox vm 
      Height          =   495
      Left            =   10800
      TabIndex        =   11
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox ux 
      Height          =   495
      Left            =   10800
      TabIndex        =   10
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox u 
      Height          =   495
      Left            =   10800
      TabIndex        =   3
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox mp 
      Height          =   495
      Left            =   10800
      TabIndex        =   2
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox m0 
      Height          =   495
      Left            =   10800
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox mt 
      Height          =   495
      Left            =   10800
      TabIndex        =   0
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Shape Shape3 
      Height          =   135
      Left            =   0
      Top             =   10200
      Width           =   135
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   15240
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      Height          =   495
      Left            =   7320
      TabIndex        =   51
      Top             =   8640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      Height          =   495
      Left            =   7320
      TabIndex        =   50
      Top             =   8040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      Height          =   495
      Left            =   7200
      TabIndex        =   49
      Top             =   7560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "150"
      Height          =   495
      Left            =   5880
      TabIndex        =   48
      Top             =   8640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "150"
      Height          =   495
      Left            =   5880
      TabIndex        =   47
      Top             =   8040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "150"
      Height          =   495
      Left            =   5760
      TabIndex        =   46
      Top             =   7560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label label6 
      Caption         =   "Label6"
      Height          =   255
      Left            =   3480
      TabIndex        =   44
      Top             =   6720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin WMPLibCtl.WindowsMediaPlayer Player2 
      Height          =   615
      Left            =   7080
      TabIndex        =   43
      Top             =   6480
      Visible         =   0   'False
      Width           =   975
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   1720
      _cy             =   1085
   End
   Begin WMPLibCtl.WindowsMediaPlayer player1 
      Height          =   735
      Left            =   480
      TabIndex        =   36
      Top             =   480
      Visible         =   0   'False
      Width           =   1815
      URL             =   "C:\Microsoft Visual Studio\VB98\misile launcher\3.mp3"
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   560
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   3201
      _cy             =   1296
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   15
      Left            =   13920
      TabIndex        =   35
      Top             =   4560
      Width           =   135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Velocity - Time Profile of Baktar Shikan Powered Sustaining Horizontal Flight."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   480
      TabIndex        =   34
      Top             =   -120
      Width           =   13215
   End
   Begin VB.Label v 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   855
      Left            =   2280
      TabIndex        =   32
      Top             =   1560
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label33 
      BackStyle       =   0  'Transparent
      Caption         =   "step2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9240
      TabIndex        =   30
      Top             =   8520
      Width           =   1335
   End
   Begin VB.Label Label31 
      BackStyle       =   0  'Transparent
      Caption         =   "Cx"
      Height          =   255
      Left            =   12240
      TabIndex        =   29
      Top             =   5160
      Width           =   735
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      Height          =   255
      Left            =   12240
      TabIndex        =   28
      Top             =   4920
      Width           =   735
   End
   Begin VB.Label Label32 
      BackStyle       =   0  'Transparent
      Caption         =   "v(0)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9240
      TabIndex        =   24
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label a 
      BackStyle       =   0  'Transparent
      Caption         =   "step3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9240
      TabIndex        =   22
      Top             =   7920
      Width           =   1455
   End
   Begin VB.Label Label30 
      BackStyle       =   0  'Transparent
      Caption         =   "m"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9240
      TabIndex        =   20
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Label Label29 
      BackStyle       =   0  'Transparent
      Caption         =   "gravity"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9240
      TabIndex        =   18
      Top             =   6720
      Width           =   1455
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "I"
      Height          =   255
      Left            =   12120
      TabIndex        =   17
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "P"
      Height          =   255
      Left            =   12120
      TabIndex        =   16
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9240
      TabIndex        =   12
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "m(p)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9240
      TabIndex        =   9
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "m(0)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9240
      TabIndex        =   8
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "u"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9240
      TabIndex        =   7
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "v(m)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9240
      TabIndex        =   6
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "u^x"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9240
      TabIndex        =   5
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "m(t)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9240
      TabIndex        =   4
      Top             =   2040
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
If Combo1.Text = "User Defined" Then
mt.Text = Text7.Text
m0.Text = Text6.Text
mp.Text = Text5.Text
u.Text = Text4.Text
ux.Text = Text3.Text
vm.Text = Text2.Text
v0.Text = Text9.Text
T.Text = Text13.Text
G.Text = Text12.Text
m.Text = Text11.Text
step3.Text = Text10.Text
step2.Text = Text8.Text
P.Text = Text17.Text
I.Text = Text16.Text
S.Text = Text15.Text
C.Text = Text14.Text
End If
If Combo1.Text = "Baktar Shikan 3 Km Range Sustainer Profile" Then
P.Text = "98.1"     'sustaining thrust'
G.Text = "9.81"     'gravity'
I.Text = "235"      'specific impulse'
T.Text = "0"        'time'
m0.Text = "10.5"    'mass at start of sustaining phase'
S.Text = "0.01131"  'reference area'
C.Text = "0.37"     'drag coefficeint'
v0.Text = "220"     'speed at start of sustaining phase'
mp.Text = "0.1"
vm.Text = "180"
End If
Player2.URL = "C:\Microsoft Visual Studio\VB98\misile launcher\UI\buttonclickrelease.wav"
End Sub

Private Sub Command1_Click()
On Error GoTo err
Timer11.Enabled = True
gl = G.Text * I.Text
If Option2.Value = True Then
mp.Text = Val(P.Text) / gl
mp.Text = mp.Text
End If

mt.Text = Val(mp.Text) * Val(T.Text)
m.Text = m0 - mt
u.Text = m.Text / m0.Text
If Option2.Value = True Then
vm.Text = (2 * P.Text / (1.23 * S.Text * C.Text)) ^ 0.5
End If
step3.Text = (1 - v0.Text / vm.Text) / (1 + v0.Text / vm.Text)
ux.Text = (2 * G.Text * I.Text) / vm.Text

d = 1 - step3.Text * u.Text ^ ux.Text
b = 1 + step3.Text * u.Text ^ ux.Text
cc = d / b
cc = cc * vm.Text
v.Caption = cc
List1.AddItem (T.Text + "         " + v.Caption)
Label47.Caption = Val(Label47.Caption) + 1
Label48.Caption = Val(Label48.Caption) + Val(v.Caption)
Text20.Text = Val(Label48.Caption) / Val(Label47.Caption)
T.Text = Val(T.Text) + Val(Text19.Text)
Exit Sub
err:
MsgBox "some or all values are not inputed that are needed to perform this operation"
End Sub

Private Sub Command10_Click()
If player1.URL = "C:\Microsoft Visual Studio\VB98\misile launcher\1.mp3" Then
player1.URL = "C:\Microsoft Visual Studio\VB98\misile launcher\2.mp3"
Exit Sub
End If
If player1.URL = "C:\Microsoft Visual Studio\VB98\misile launcher\2.mp3" Then
player1.URL = "C:\Microsoft Visual Studio\VB98\misile launcher\3.mp3"
Exit Sub
End If
If player1.URL = "C:\Microsoft Visual Studio\VB98\misile launcher\3.mp3" Then
player1.URL = "C:\Microsoft Visual Studio\VB98\misile launcher\1.mp3"
Exit Sub
End If
End Sub

Private Sub Command11_Click()
On Error GoTo err
Adodc1.Recordset.AddNew
Exit Sub
err:
MsgBox "error check with Ahmed Adnan.restart project"
End Sub

Private Sub Command12_Click()
On Error GoTo err
Adodc1.Recordset.Save
Exit Sub
err:
MsgBox "error check with Ahmed Adnan.restart project"
End Sub

Private Sub Command13_Click()
On Error GoTo err
Adodc1.Recordset.Delete
Exit Sub
err:
MsgBox "error check with Ahmed Adnan.restart project"
End Sub

Private Sub Command14_Click()
Frame8.Visible = False
    With Form1
.mt.Text = .Text7.Text
.m0.Text = .Text6.Text
.mp.Text = .Text5.Text
.u.Text = .Text4.Text
.ux.Text = .Text3.Text
.vm.Text = .Text2.Text
.v0.Text = .Text9.Text
.T.Text = .Text13.Text
.G.Text = .Text12.Text
.m.Text = .Text11.Text
.step3.Text = .Text10.Text
.step2.Text = .Text8.Text
.P.Text = .Text17.Text
.I.Text = .Text16.Text
.S.Text = .Text15.Text
.C.Text = .Text14.Text
End With
End Sub

Private Sub Command15_Click()
If Adodc1.Recordset.EOF = False Then
Adodc1.Recordset.MoveNext
End If
If Label50.Caption = Empty Then
Adodc1.Recordset.MovePrevious
End If
End Sub

Private Sub Command16_Click()
If Adodc1.Recordset.BOF = False Then
Adodc1.Recordset.MovePrevious
End If
If Label50.Caption = Empty Then
Adodc1.Recordset.MoveNext
End If
End Sub

Private Sub Command17_Click()
MsgBox "if you recieve a message like this 'some or all values are not inputed that are needed to perform this operation'then it means that ###you have either inputted a value incorrectly###  ###you have not inputted one or more values needed for this operation###  ###the calculation has reached a point after which further calculations cant be calculated given the formulae used###  ###you have chosen the wrong settings and you should goto to 'settings' and switch to 'automatically'   "
MsgBox "if you receive a message like this 'error check with Ahmed Adnan.restart project' then it means at the program has not been developed that much so you should be more carefull and not press buttons at random but in sequence so that all operations work in unisom/rythem or whatever the word "
End Sub

Private Sub Command2_Click()
Timer4.Enabled = True
P.Visible = False     'sustaining thrust'
G.Visible = False   'gravity'
I.Visible = False      'specific impulse'
T.Visible = True       'time'
m0.Visible = False    'mass at start of sustaining phase'
S.Visible = False  'reference area'
C.Visible = False     'drag coefficeint'
v0.Visible = False     'speed at start of sustaining phase'
vm.Visible = False
mp.Visible = False
mt.Visible = False
u.Visible = False
ux.Visible = False
m.Visible = False
step3.Visible = False
step2.Visible = False
T.Visible = False
Command3.Visible = False
Frame2.Visible = False
Timer5.Enabled = True
Frame2.Visible = False
SSTab1.Visible = True
Frame4.Visible = True

End Sub

Private Sub Command3_Click()

Unload Form1
Load Form1
Form1.Show
End Sub

Private Sub Command4_Click()
T.Visible = True
Frame4.Visible = False
Frame2.Visible = True
Command3.Visible = True
Timer3.Enabled = True
mp.Visible = True
vm.Visible = True
P.Visible = True
G.Visible = True
I.Visible = True
T.Visible = True
m0.Visible = True
S.Visible = True
C.Visible = True
v0.Visible = True
vm.Visible = True
mp.Visible = True
mt.Visible = True
u.Visible = True
ux.Visible = True
m.Visible = True
step3.Visible = True
step2.Visible = True
mt.Text = Text7.Text
m0.Text = Text6.Text
mp.Text = Text5.Text
u.Text = Text4.Text
ux.Text = Text3.Text
vm.Text = Text2.Text
v0.Text = Text9.Text
T.Text = Text13.Text
G.Text = Text12.Text
m.Text = Text11.Text
step3.Text = Text10.Text
step2.Text = Text8.Text
P.Text = Text17.Text
I.Text = Text16.Text
S.Text = Text15.Text
C.Text = Text14.Text
Label10.Caption = "100"
Label11.Caption = "100"
Label12.Caption = "100"
X = Label10.Caption
Y = Label11.Caption
z = Label12.Caption
Frame1.BackColor = RGB(X, Y, z)
Frame4.BackColor = RGB(X, Y, z)
Frame5.BackColor = RGB(X, Y, z)
SSTab1.Visible = False
Player2.URL = "C:\Microsoft Visual Studio\VB98\misile launcher\UI\buttonclick.wav"
End Sub

Private Sub Command6_Click()
Timer7.Enabled = True
Timer10.Enabled = True
End Sub

Private Sub Command7_Click()
Timer7.Enabled = False
End Sub

Private Sub Command8_Click()
If Text1.Text = Text1.Tag Then
Frame7.Visible = True
Timer8.Enabled = True
Timer9.Enabled = False
Command9.Visible = True
Player2.URL = "C:\Microsoft Visual Studio\VB98\misile launcher\UI\button1.wav"
Command8.Enabled = False
Command11.Visible = True
Command12.Visible = True
Command13.Visible = True
Adodc1.Visible = True
Else
Player2.URL = "C:\Microsoft Visual Studio\VB98\misile launcher\UI\combine_button_locked.wav"
End If
End Sub

Private Sub Command9_Click()
Player2.URL = "C:\Microsoft Visual Studio\VB98\misile launcher\UI\buttonclick.wav"
Frame7.Visible = False
Timer8.Enabled = False
Timer9.Enabled = True
Text1.Text = Empty
Command9.Visible = False
Command8.Enabled = True
Command11.Visible = False
Command12.Visible = False
Command13.Visible = False
Adodc1.Visible = False
End Sub

Private Sub Form_Load()
player1.settings.volume = 20


End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
label6.Caption = "0"
End Sub

Private Sub Frame4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
label6.Caption = "0"
End Sub

Private Sub Option1_Click()
Option1.Value = True
Option2.Value = False
Shape1.BorderColor = &H8000000D
Shape2.Visible = False
Shape1.Visible = True
Player2.URL = "C:\Microsoft Visual Studio\VB98\misile launcher\UI\buttonclickrelease.wav"
End Sub

Private Sub Option1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If label6.Caption <> "1" Then
Player2.URL = "C:\Microsoft Visual Studio\VB98\misile launcher\UI\buttonrollover.wav"
label6.Caption = "1"
End If
End Sub

Private Sub Option2_Click()
Option2.Value = True
Option1.Value = False
Shape2.Visible = True
Shape1.Visible = False
Shape2.BorderColor = &H8000000D
Player2.URL = "C:\Microsoft Visual Studio\VB98\misile launcher\UI\buttonclickrelease.wav"
End Sub

Private Sub WindowsMediaPlayer1_EndOfStream(ByVal Result As Long)

End Sub

Private Sub WindowsMediaPlayer1_OpenStateChange(ByVal NewState As Long)

End Sub

Private Sub Option2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If label6.Caption <> "1" Then
Player2.URL = "C:\Microsoft Visual Studio\VB98\misile launcher\UI\buttonrollover.wav"
label6.Caption = "1"
End If
End Sub

Private Sub player1_EndOfStream(ByVal Result As Long)
MsgBox "1"
If player1.URL = "C:\Microsoft Visual Studio\VB98\misile launcher\1.mp3" Then
player1.URL = "C:\Microsoft Visual Studio\VB98\misile launcher\2.mp3"
Exit Sub
End If
If player1.URL = "C:\Microsoft Visual Studio\VB98\misile launcher\2.mp3" Then
player1.URL = "C:\Microsoft Visual Studio\VB98\misile launcher\3.mp3"
Exit Sub
End If
If player1.URL = "C:\Microsoft Visual Studio\VB98\misile launcher\3.mp3" Then
player1.URL = "C:\Microsoft Visual Studio\VB98\misile launcher\1.mp3"
Exit Sub
End If
MsgBox ""
End Sub

Private Sub player_OpenStateChange(ByVal NewState As Long)

End Sub

Private Sub Timer1_Timer()
If Option2.Value = True Then
mp.Enabled = False
vm.Enabled = False
P.Enabled = True
I.Enabled = True
S.Enabled = True
C.Enabled = True
Else
mp.Enabled = True
vm.Enabled = True
P.Enabled = False
I.Enabled = False
S.Enabled = False
C.Enabled = False
End If
If Combo1.Text = "Baktar Shikan 3 Km Range Sustainer Profile" Then
P.Enabled = False     'sustaining thrust'
I.Enabled = False      'specific impulse'
T.Enabled = True       'time'
m0.Enabled = False    'mass at start of sustaining phase'
S.Enabled = False  'reference area'
C.Enabled = False     'drag coefficeint'
v0.Enabled = False     'speed at start of sustaining phase'


mt.Enabled = False
u.Enabled = False
ux.Enabled = False
m.Enabled = False
step3.Enabled = False
step2.Enabled = False
Else
P.Enabled = True
I.Enabled = True
T.Enabled = True
m0.Enabled = True
S.Enabled = True
C.Enabled = True
v0.Enabled = True
mt.Enabled = True
u.Enabled = True
ux.Enabled = True
m.Enabled = True
step3.Enabled = True
step2.Enabled = True
End If
End Sub

Private Sub Timer10_Timer()
If T.Text > Text18.Text Then
Timer7.Enabled = False
Timer10.Enabled = False
End If
End Sub

Private Sub Timer11_Timer()
Shape3.Left = Shape3.Left + v.Caption
End Sub

Private Sub Timer3_Timer()
If Label7.Caption <= 200 Then
Label7.Caption = Val(Label7.Caption) + 5
Label8.Caption = Val(Label8.Caption) + 5
Label9.Caption = Val(Label9.Caption) + 5
X = Label7.Caption
Y = Label8.Caption
z = Label9.Caption
Form1.BackColor = RGB(X, Y, z)
Else
Timer3.Enabled = False
End If
End Sub

Private Sub Timer4_Timer()
If Label7.Caption >= 100 Then
Label7.Caption = Val(Label7.Caption) - 5
Label8.Caption = Val(Label8.Caption) - 5
Label9.Caption = Val(Label9.Caption) - 5
X = Label7.Caption
Y = Label8.Caption
z = Label9.Caption
Form1.BackColor = RGB(X, Y, z)
Else
Timer4.Enabled = False
End If
End Sub

Private Sub Timer5_Timer()
If Label10.Caption < 235 Then
Label10.Caption = Val(Label10.Caption) + 5
Label11.Caption = Val(Label11.Caption) + 5
Label12.Caption = Val(Label12.Caption) + 5
X = Label10.Caption
Y = Label11.Caption
z = Label12.Caption
Frame1.BackColor = RGB(X, Y, z)
Frame4.BackColor = RGB(X, Y, z)
Frame5.BackColor = RGB(X, Y, z)
Else
Timer5.Enabled = False
End If
End Sub

Private Sub Timer7_Timer()
On Error GoTo err
'If Text19.Text < Text20.Text Then
gl = G.Text * I.Text
If Option2.Value = True Then
mp.Text = Val(P.Text) / gl
mp.Text = mp.Text
End If

mt.Text = Val(mp.Text) * Val(T.Text)
m.Text = m0 - mt
u.Text = m.Text / m0.Text
If Option2.Value = True Then
vm.Text = (2 * P.Text / (1.23 * S.Text * C.Text)) ^ 0.5
End If
step3.Text = (1 - v0.Text / vm.Text) / (1 + v0.Text / vm.Text)
ux.Text = (2 * G.Text * I.Text) / vm.Text

d = 1 - step3.Text * u.Text ^ ux.Text
b = 1 + step3.Text * u.Text ^ ux.Text
cc = d / b
cc = cc * vm.Text
v.Caption = cc
List1.AddItem (T.Text)
T.Text = T.Text + 1
List2.AddItem (v.Caption)
Shape3.Top = 9000 - T.Text
Shape3.Left = v.Caption * 10
'End If
Exit Sub
err:
MsgBox "some or all values are not inputed that are needed to perform this operation"
Timer7.Enabled = False
End Sub

Private Sub Timer8_Timer()
If SSTab1.Height < 7300 Then
SSTab1.Height = SSTab1.Height + 35
End If
If SSTab1.Width < 8300 Then
SSTab1.Width = SSTab1.Width + 35
End If
End Sub

Private Sub Timer9_Timer()
If SSTab1.Height > 4575 Then
SSTab1.Height = SSTab1.Height - 100
End If
If SSTab1.Width > 5295 Then
SSTab1.Width = SSTab1.Width - 100
End If
End Sub
