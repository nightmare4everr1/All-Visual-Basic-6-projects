VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Barclay's Premiere Bank System"
   ClientHeight    =   11010
   ClientLeft      =   105
   ClientTop       =   0
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command12 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   24
         Charset         =   1
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   12360
      TabIndex        =   117
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox Text18 
      DataField       =   "deposits"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   2280
      TabIndex        =   94
      Text            =   "Text18"
      Top             =   2760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text17 
      DataField       =   "withdrawels"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   2280
      TabIndex        =   93
      Text            =   "Text17"
      Top             =   3120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   21.75
         Charset         =   1
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   12360
      TabIndex        =   88
      Top             =   240
      Width           =   2295
   End
   Begin VB.TextBox Text13 
      DataField       =   "cashinpack"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1560
      TabIndex        =   79
      Text            =   "Text13"
      Top             =   9840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3120
      Top             =   10440
   End
   Begin VB.TextBox Text12 
      DataField       =   "packinterest"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1560
      TabIndex        =   77
      Text            =   "Text12"
      Top             =   9480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7575
      Left            =   2880
      TabIndex        =   28
      Top             =   2400
      Visible         =   0   'False
      Width           =   8565
      _ExtentX        =   15108
      _ExtentY        =   13361
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabHeight       =   520
      TabCaption(0)   =   "general details"
      TabPicture(0)   =   "Form1.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Timer6"
      Tab(0).Control(1)=   "Timer5"
      Tab(0).Control(2)=   "Command6"
      Tab(0).Control(3)=   "Timer1"
      Tab(0).Control(4)=   "Command8"
      Tab(0).Control(5)=   "Timer2"
      Tab(0).Control(6)=   "Command7"
      Tab(0).Control(7)=   "Text10"
      Tab(0).Control(8)=   "Command9"
      Tab(0).Control(9)=   "Timer3"
      Tab(0).Control(10)=   "Label70"
      Tab(0).Control(11)=   "Label60"
      Tab(0).Control(12)=   "Label59"
      Tab(0).Control(13)=   "Label58"
      Tab(0).Control(14)=   "Label57"
      Tab(0).Control(15)=   "Label56"
      Tab(0).Control(16)=   "Label55"
      Tab(0).Control(17)=   "Label54"
      Tab(0).Control(18)=   "Label53"
      Tab(0).Control(19)=   "Line1"
      Tab(0).Control(20)=   "Label14"
      Tab(0).Control(21)=   "Label15"
      Tab(0).Control(22)=   "Label16"
      Tab(0).Control(23)=   "Label17"
      Tab(0).Control(24)=   "Label18"
      Tab(0).Control(25)=   "Label19"
      Tab(0).Control(26)=   "Label22"
      Tab(0).Control(27)=   "Label23"
      Tab(0).Control(28)=   "Label24"
      Tab(0).Control(29)=   "Label25"
      Tab(0).Control(30)=   "Label26"
      Tab(0).Control(31)=   "Label27"
      Tab(0).Control(32)=   "Label28"
      Tab(0).Control(33)=   "Label29"
      Tab(0).Control(34)=   "Label30"
      Tab(0).Control(35)=   "Label20"
      Tab(0).Control(36)=   "Label21"
      Tab(0).Control(37)=   "Label31"
      Tab(0).Control(38)=   "Label32"
      Tab(0).Control(39)=   "Label33"
      Tab(0).ControlCount=   40
      TabCaption(1)   =   "package details"
      TabPicture(1)   =   "Form1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Timer9"
      Tab(1).Control(1)=   "Text11"
      Tab(1).Control(2)=   "Command10"
      Tab(1).Control(3)=   "Command11"
      Tab(1).Control(4)=   "Combo2"
      Tab(1).Control(5)=   "Command13"
      Tab(1).Control(6)=   "Label52"
      Tab(1).Control(7)=   "Label51"
      Tab(1).Control(8)=   "Label34"
      Tab(1).Control(9)=   "Label35"
      Tab(1).Control(10)=   "Label36"
      Tab(1).Control(11)=   "Label37"
      Tab(1).Control(12)=   "Label38"
      Tab(1).Control(13)=   "Label39"
      Tab(1).Control(14)=   "Label40"
      Tab(1).Control(15)=   "Label41"
      Tab(1).Control(16)=   "Label44"
      Tab(1).Control(17)=   "Label45"
      Tab(1).Control(18)=   "Label46"
      Tab(1).Control(19)=   "Label47"
      Tab(1).Control(20)=   "Label48"
      Tab(1).Control(21)=   "Label49"
      Tab(1).Control(22)=   "Label50"
      Tab(1).ControlCount=   23
      TabCaption(2)   =   "transactions"
      TabPicture(2)   =   "Form1.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Command17"
      Tab(2).Control(1)=   "Timer11"
      Tab(2).Control(2)=   "Timer10"
      Tab(2).Control(3)=   "List1"
      Tab(2).Control(4)=   "List2"
      Tab(2).Control(5)=   "Label68"
      Tab(2).Control(6)=   "Label67"
      Tab(2).Control(7)=   "Label66"
      Tab(2).Control(8)=   "Label65"
      Tab(2).Control(9)=   "Label64"
      Tab(2).ControlCount=   10
      TabCaption(3)   =   "account handling"
      TabPicture(3)   =   "Form1.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Label73"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label74"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label75"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label76"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Label77"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Label78"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "Label79"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "Label80"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "Line4"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "Line5"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "Label69"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "Text16"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "Text19"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).Control(13)=   "Command19"
      Tab(3).Control(13).Enabled=   0   'False
      Tab(3).Control(14)=   "Text21"
      Tab(3).Control(14).Enabled=   0   'False
      Tab(3).Control(15)=   "Command20"
      Tab(3).Control(15).Enabled=   0   'False
      Tab(3).Control(16)=   "Command14"
      Tab(3).Control(16).Enabled=   0   'False
      Tab(3).ControlCount=   17
      Begin VB.CommandButton Command14 
         Caption         =   "Quit"
         Height          =   495
         Left            =   3600
         TabIndex        =   120
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton Command20 
         Caption         =   "Apply changes"
         Height          =   495
         Left            =   720
         TabIndex        =   118
         Top             =   4680
         Width           =   2175
      End
      Begin VB.CommandButton Command17 
         Caption         =   "refresh data"
         Height          =   495
         Left            =   -71040
         TabIndex        =   115
         Top             =   3720
         Width           =   1215
      End
      Begin VB.TextBox Text21 
         DataField       =   "account reference"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   1440
         TabIndex        =   113
         Text            =   "Text21"
         Top             =   4200
         Width           =   1575
      End
      Begin VB.CommandButton Command19 
         Caption         =   "change pin code"
         Height          =   495
         Left            =   1440
         TabIndex        =   112
         Top             =   3600
         Width           =   1575
      End
      Begin VB.TextBox Text19 
         DataField       =   "address"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   1440
         TabIndex        =   111
         Text            =   "Text19"
         Top             =   2520
         Width           =   1575
      End
      Begin VB.TextBox Text16 
         DataField       =   "name"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   1440
         TabIndex        =   110
         Text            =   "Text16"
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Timer Timer11 
         Interval        =   1
         Left            =   -73200
         Top             =   5520
      End
      Begin VB.Timer Timer10 
         Interval        =   1
         Left            =   -71640
         Top             =   5520
      End
      Begin VB.Timer Timer7 
         Interval        =   1
         Left            =   7680
         Top             =   -1560
      End
      Begin VB.ListBox List1 
         Height          =   1815
         ItemData        =   "Form1.frx":0070
         Left            =   -73200
         List            =   "Form1.frx":0072
         TabIndex        =   96
         Top             =   3360
         Width           =   1455
      End
      Begin VB.ListBox List2 
         Height          =   1815
         ItemData        =   "Form1.frx":0074
         Left            =   -69360
         List            =   "Form1.frx":0076
         TabIndex        =   95
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Timer Timer9 
         Interval        =   100
         Left            =   -72360
         Top             =   1260
      End
      Begin VB.Timer Timer6 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   -70440
         Top             =   6180
      End
      Begin VB.Timer Timer5 
         Interval        =   1
         Left            =   -69120
         Top             =   5460
      End
      Begin VB.TextBox Text11 
         Height          =   525
         Left            =   -68640
         TabIndex        =   73
         Text            =   "0"
         Top             =   1620
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Deactivate package"
         Height          =   495
         Left            =   -71760
         TabIndex        =   57
         Top             =   1620
         Width           =   1575
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Activate package"
         Height          =   495
         Left            =   -71760
         TabIndex        =   56
         Top             =   960
         Width           =   1575
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "Form1.frx":0078
         Left            =   -68640
         List            =   "Form1.frx":0088
         TabIndex        =   55
         Top             =   900
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton Command13 
         Caption         =   "click to see terms and conditions"
         Height          =   495
         Left            =   -68880
         TabIndex        =   54
         Top             =   6660
         Width           =   1935
      End
      Begin VB.CommandButton Command6 
         Caption         =   "reset predictions"
         Height          =   495
         Left            =   -69840
         TabIndex        =   53
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   -70560
         Top             =   2340
      End
      Begin VB.CommandButton Command8 
         Caption         =   "time travel"
         Height          =   495
         Left            =   -71040
         TabIndex        =   32
         Top             =   4680
         Width           =   975
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   -70080
         Top             =   2340
      End
      Begin VB.CommandButton Command7 
         Caption         =   "predict balance after next session"
         Height          =   495
         Left            =   -74880
         TabIndex        =   31
         Top             =   3060
         Width           =   1695
      End
      Begin VB.TextBox Text10 
         Height          =   495
         Left            =   -72600
         TabIndex        =   30
         Top             =   3660
         Width           =   1095
      End
      Begin VB.CommandButton Command9 
         Caption         =   "predict sessions needed"
         Height          =   495
         Left            =   -71400
         TabIndex        =   29
         Top             =   3660
         Width           =   1455
      End
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   -69600
         Top             =   2340
      End
      Begin VB.Label Label70 
         Caption         =   "predictor"
         ForeColor       =   &H00C000C0&
         Height          =   375
         Left            =   -72960
         TabIndex        =   116
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label Label69 
         Alignment       =   2  'Center
         Caption         =   "Label69"
         DataField       =   "amount"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   114
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Line Line5 
         X1              =   0
         X2              =   3480
         Y1              =   5280
         Y2              =   5280
      End
      Begin VB.Line Line4 
         X1              =   3480
         X2              =   3480
         Y1              =   600
         Y2              =   5280
      End
      Begin VB.Label Label80 
         Caption         =   "Name"
         Height          =   495
         Left            =   360
         TabIndex        =   109
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label79 
         Caption         =   "Address"
         Height          =   495
         Left            =   360
         TabIndex        =   108
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label78 
         Caption         =   "Account ID"
         Height          =   375
         Left            =   360
         TabIndex        =   107
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label77 
         Caption         =   "Balance"
         Height          =   495
         Left            =   360
         TabIndex        =   106
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label76 
         Caption         =   "Pin code"
         Height          =   495
         Left            =   360
         TabIndex        =   105
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label Label75 
         Caption         =   "Account reference"
         Height          =   495
         Left            =   240
         TabIndex        =   104
         Top             =   4200
         Width           =   1095
      End
      Begin VB.Label Label74 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         DataField       =   "roll"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "HalfLife2"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1440
         TabIndex        =   103
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label73 
         Caption         =   "Your account info"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   495
         Left            =   480
         TabIndex        =   102
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label Label68 
         BackStyle       =   0  'Transparent
         Caption         =   "name"
         DataField       =   "name"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   735
         Left            =   -71040
         TabIndex        =   101
         Top             =   1440
         Width           =   4215
      End
      Begin VB.Label Label67 
         BackStyle       =   0  'Transparent
         Caption         =   "Your transaction history"
         BeginProperty Font 
            Name            =   "Marking Pen"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -73560
         TabIndex        =   100
         Top             =   840
         Width           =   4815
      End
      Begin VB.Label Label66 
         Caption         =   "Withdrawels"
         BeginProperty Font 
            Name            =   "Marking Pen"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   495
         Left            =   -73320
         TabIndex        =   99
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Label Label65 
         Caption         =   "Deposits"
         BeginProperty Font 
            Name            =   "Marking Pen"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   495
         Left            =   -69240
         TabIndex        =   98
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label64 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -73440
         TabIndex        =   97
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label60 
         Caption         =   "Label60"
         Height          =   375
         Left            =   -70080
         TabIndex        =   87
         Top             =   5580
         Width           =   855
      End
      Begin VB.Label Label59 
         Caption         =   "Label59"
         DataField       =   "cashinpack"
         DataSource      =   "Adodc1"
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   -71400
         TabIndex        =   86
         Top             =   5940
         Width           =   1335
      End
      Begin VB.Label Label58 
         Caption         =   "amount deposited"
         Height          =   375
         Left            =   -72720
         TabIndex        =   85
         Top             =   5940
         Width           =   1335
      End
      Begin VB.Label Label57 
         Caption         =   "Label57"
         DataField       =   "packinterest"
         DataSource      =   "Adodc1"
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   -73680
         TabIndex        =   84
         Top             =   5940
         Width           =   615
      End
      Begin VB.Label Label56 
         Caption         =   "interest rate"
         Height          =   375
         Left            =   -74880
         TabIndex        =   83
         Top             =   5940
         Width           =   1215
      End
      Begin VB.Label Label55 
         Caption         =   "after one session ="
         Height          =   255
         Left            =   -71520
         TabIndex        =   82
         Top             =   5580
         Width           =   1455
      End
      Begin VB.Label Label54 
         Caption         =   "Label54"
         DataField       =   "packages"
         DataSource      =   "Adodc1"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -72720
         TabIndex        =   81
         Top             =   5580
         Width           =   1335
      End
      Begin VB.Label Label53 
         Caption         =   "predicted cash of package "
         Height          =   375
         Left            =   -74880
         TabIndex        =   80
         Top             =   5580
         Width           =   2295
      End
      Begin VB.Line Line1 
         X1              =   -75000
         X2              =   -66600
         Y1              =   5340
         Y2              =   5340
      End
      Begin VB.Label Label52 
         Caption         =   "Label52"
         DataField       =   "amount"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   -68640
         TabIndex        =   75
         Top             =   1260
         Width           =   1575
      End
      Begin VB.Label Label51 
         Caption         =   "Current balance"
         Height          =   375
         Left            =   -69960
         TabIndex        =   74
         Top             =   1260
         Width           =   1215
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         Caption         =   "3 year pack"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -73920
         TabIndex        =   72
         Top             =   5220
         Width           =   1215
      End
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0FF&
         Caption         =   "the conditions are the same as of the previous pack,but you get 22%extra surplus"
         Height          =   495
         Left            =   -72120
         TabIndex        =   71
         Top             =   5340
         Width           =   3135
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "5 year pack"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   -73920
         TabIndex        =   70
         Top             =   6060
         Width           =   1215
      End
      Begin VB.Label Label37 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "same except you get 30% extra"
         Height          =   495
         Left            =   -72120
         TabIndex        =   69
         Top             =   6120
         Width           =   3135
      End
      Begin VB.Label Label38 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         Caption         =   "ten year pack"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -73920
         TabIndex        =   68
         Top             =   6900
         Width           =   1215
      End
      Begin VB.Label Label39 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0FF&
         Caption         =   "same except you get 50% extra"
         Height          =   495
         Left            =   -72120
         TabIndex        =   67
         Top             =   6900
         Width           =   3135
      End
      Begin VB.Label Label40 
         Caption         =   "Deposit amount"
         Height          =   255
         Left            =   -69960
         TabIndex        =   66
         Top             =   1740
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label41 
         Caption         =   "Package type"
         Height          =   495
         Left            =   -69960
         TabIndex        =   65
         Top             =   900
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label44 
         Alignment       =   2  'Center
         Caption         =   "Your current package is"
         Height          =   495
         Left            =   -74760
         TabIndex        =   64
         Top             =   900
         Width           =   1215
      End
      Begin VB.Label Label45 
         Alignment       =   2  'Center
         Caption         =   "Label45"
         DataField       =   "packages"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   -73320
         TabIndex        =   63
         Top             =   900
         Width           =   1215
      End
      Begin VB.Label Label46 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         Caption         =   "Package Types"
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
         Left            =   -72600
         TabIndex        =   62
         Top             =   2340
         Width           =   3855
      End
      Begin VB.Label Label47 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         Caption         =   "None(default)"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   -73920
         TabIndex        =   61
         Top             =   3420
         Width           =   1215
      End
      Begin VB.Label Label48 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0FF&
         Caption         =   "this is the default package.at the end of every session you get 10% surplus cash of your deposited amount"
         Height          =   495
         Left            =   -72120
         TabIndex        =   60
         Top             =   3420
         Width           =   3855
      End
      Begin VB.Label Label49 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "1 year pack"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   -73920
         TabIndex        =   59
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Label Label50 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   $"Form1.frx":00C1
         Height          =   855
         Left            =   -72120
         TabIndex        =   58
         Top             =   4140
         Width           =   5295
      End
      Begin VB.Label Label14 
         DataField       =   "amount"
         DataSource      =   "Adodc1"
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   -73560
         TabIndex        =   52
         Top             =   900
         Width           =   1935
      End
      Begin VB.Label Label15 
         DataSource      =   "Adodc1"
         Height          =   735
         Left            =   -70320
         TabIndex        =   51
         Top             =   1500
         Width           =   2055
      End
      Begin VB.Label Label16 
         DataField       =   "interest"
         DataSource      =   "Adodc1"
         ForeColor       =   &H8000000D&
         Height          =   615
         Left            =   -70320
         TabIndex        =   50
         Top             =   780
         Width           =   2175
      End
      Begin VB.Label Label17 
         DataField       =   "packages"
         DataSource      =   "Adodc1"
         ForeColor       =   &H8000000D&
         Height          =   735
         Left            =   -73560
         TabIndex        =   49
         Top             =   1380
         Width           =   1695
      End
      Begin VB.Label Label18 
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   -72600
         TabIndex        =   48
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label19 
         Caption         =   "time passed since cash deposited"
         Height          =   495
         Left            =   -74760
         TabIndex        =   47
         Top             =   4500
         Width           =   1695
      End
      Begin VB.Label Label22 
         Caption         =   "day number"
         Height          =   255
         Left            =   -72840
         TabIndex        =   46
         Top             =   4500
         Width           =   975
      End
      Begin VB.Label Label23 
         Caption         =   "months"
         Height          =   375
         Left            =   -72960
         TabIndex        =   45
         Top             =   4920
         Width           =   1095
      End
      Begin VB.Label Label24 
         Caption         =   "25"
         Height          =   255
         Left            =   -71760
         TabIndex        =   44
         Top             =   4500
         Width           =   255
      End
      Begin VB.Label Label25 
         Caption         =   "0"
         Height          =   255
         Left            =   -71760
         TabIndex        =   43
         Top             =   4980
         Width           =   255
      End
      Begin VB.Label Label26 
         Caption         =   "period of deposit"
         Height          =   495
         Left            =   -71640
         TabIndex        =   42
         Top             =   1500
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label27 
         Caption         =   "interest rate"
         Height          =   495
         Left            =   -71640
         TabIndex        =   41
         Top             =   780
         Width           =   1215
      End
      Begin VB.Label Label28 
         Caption         =   "your current account"
         Height          =   495
         Left            =   -74880
         TabIndex        =   40
         Top             =   780
         Width           =   1215
      End
      Begin VB.Label Label29 
         Caption         =   "package"
         Height          =   615
         Left            =   -74880
         TabIndex        =   39
         Top             =   1500
         Width           =   1215
      End
      Begin VB.Label Label30 
         Caption         =   "predictid surplus after 6 months(one session)"
         Height          =   615
         Left            =   -74880
         TabIndex        =   38
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Label Label20 
         Caption         =   "session number"
         Height          =   255
         Left            =   -71640
         TabIndex        =   37
         Top             =   3180
         Width           =   1215
      End
      Begin VB.Label Label21 
         ForeColor       =   &H00FF00FF&
         Height          =   375
         Left            =   -70440
         TabIndex        =   36
         Top             =   3180
         Width           =   975
      End
      Begin VB.Label Label31 
         Caption         =   "predict how many sessions are needed to achive the target of"
         Height          =   495
         Left            =   -74880
         TabIndex        =   35
         Top             =   3660
         Width           =   2295
      End
      Begin VB.Label Label32 
         Caption         =   "years needed"
         Height          =   375
         Left            =   -69480
         TabIndex        =   34
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label33 
         ForeColor       =   &H00FF00FF&
         Height          =   375
         Left            =   -68400
         TabIndex        =   33
         Top             =   3120
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Easy Transaction"
      Height          =   5775
      Left            =   0
      TabIndex        =   16
      Top             =   1320
      Width           =   2655
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   1680
         TabIndex        =   90
         Top             =   4560
         Width           =   855
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Deposit additional amount"
         Enabled         =   0   'False
         Height          =   495
         Left            =   120
         TabIndex        =   89
         Top             =   4440
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "_"
         TabIndex        =   22
         Top             =   720
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ENTER PIN"
         Height          =   495
         Left            =   480
         TabIndex        =   21
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Change pin code"
         Enabled         =   0   'False
         Height          =   495
         Left            =   240
         TabIndex        =   20
         Top             =   3120
         Width           =   2055
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Withdraw cash ammount "
         Enabled         =   0   'False
         Height          =   495
         Left            =   120
         TabIndex        =   19
         Top             =   3720
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         Caption         =   "End transaction"
         Enabled         =   0   'False
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   5160
         Width           =   2415
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   1680
         TabIndex        =   17
         Top             =   3840
         Width           =   855
      End
      Begin VB.Label Label63 
         Caption         =   "Label63"
         DataField       =   "roll"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   1440
         TabIndex        =   92
         Top             =   2040
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label62 
         Caption         =   "Account ID"
         Height          =   375
         Left            =   120
         TabIndex        =   91
         Top             =   2040
         Width           =   975
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   2640
         Y1              =   5040
         Y2              =   5040
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   2640
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Label Label2 
         Caption         =   " Balance"
         Height          =   495
         Left            =   120
         TabIndex        =   27
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "label3"
         DataField       =   "amount"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   960
         TabIndex        =   26
         Top             =   2880
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Account"
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "label5"
         DataField       =   "name"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   1200
         TabIndex        =   24
         Top             =   2520
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "Credit card machine"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.TextBox Text9 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   5520
      PasswordChar    =   "_"
      TabIndex        =   15
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "access your account"
      Height          =   495
      Left            =   3960
      TabIndex        =   13
      Top             =   1800
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "packages"
      DataSource      =   "Adodc1"
      Height          =   315
      ItemData        =   "Form1.frx":01D0
      Left            =   1560
      List            =   "Form1.frx":01E3
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   9120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      DataField       =   "interest"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Text            =   "Text6"
      Top             =   8760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      DataField       =   "amount"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Text            =   "Text5"
      Top             =   8400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      DataField       =   "roll"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Text            =   "Text4"
      Top             =   8040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      DataField       =   "pin"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   7680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      DataField       =   "name"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   7320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   240
      Top             =   10200
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Microsoft Visual Studio\VB98\bank project\bank.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Microsoft Visual Studio\VB98\bank project\bank.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "bank"
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
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   615
      Left            =   360
      TabIndex        =   119
      Top             =   360
      Visible         =   0   'False
      Width           =   1575
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   3
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   "C:\Microsoft Visual Studio\VB98\bank project\sounds\buttons\coinslot1"
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   0
      WindowlessVideo =   0   'False
   End
   Begin VB.Label Label43 
      Caption         =   "deposited amount in pack"
      DragMode        =   1  'Automatic
      Height          =   375
      Left            =   240
      TabIndex        =   78
      Top             =   9720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label42 
      Caption         =   "package interest"
      DragMode        =   1  'Automatic
      Height          =   375
      Left            =   240
      TabIndex        =   76
      Top             =   9480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "name"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   855
      Left            =   2760
      TabIndex        =   14
      Top             =   960
      Visible         =   0   'False
      Width           =   9495
   End
   Begin VB.Label Label11 
      Caption         =   "package"
      DragMode        =   1  'Automatic
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   9120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "interest"
      DragMode        =   1  'Automatic
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   8760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "ammount"
      DragMode        =   1  'Automatic
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   8400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "roll"
      DragMode        =   1  'Automatic
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   8040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "pin"
      DragMode        =   1  'Automatic
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   7680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "name"
      DragMode        =   1  'Automatic
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   7320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Barclay's Premiere Bank"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1095
      Left            =   2280
      TabIndex        =   0
      Top             =   0
      Width           =   9135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Combo1_Change()
On Error GoTo err:
If Adodc1.Recordset.Fields(5) = "1 year pack" Then
Adodc1.Recordset.Fields(6) = "13"
End If
If Adodc1.Recordset.Fields(5) = "3 year pack" Then
Adodc1.Recordset.Fields(6) = "22"
End If
If Adodc1.Recordset.Fields(5) = "5 year pack" Then
Adodc1.Recordset.Fields(6) = "30"
End If
If Adodc1.Recordset.Fields(5) = "10 year pack" Then
Adodc1.Recordset.Fields(6) = "50"
End If
If Adodc1.Recordset.Fields(5) = "none" Then
Adodc1.Recordset.Fields(6) = "0"
End If
Exit Sub
err:
MsgBox "error! combo1"
End Sub




Private Sub Combo1_Click()
If Adodc1.Recordset.Fields(5) = "1 year pack" Then Adodc1.Recordset.Fields(6) = "13"
If Adodc1.Recordset.Fields(5) = "3 year pack" Then Adodc1.Recordset.Fields(6) = "22"

If Adodc1.Recordset.Fields(5) = "5 year pack" Then Adodc1.Recordset.Fields(6) = "30"

If Adodc1.Recordset.Fields(5) = "10 year pack" Then Adodc1.Recordset.Fields(6) = "50"

If Adodc1.Recordset.Fields(5) = "none" Then Adodc1.Recordset.Fields(6) = "0"
End Sub

Private Sub Command1_Click()
Command5.Enabled = False
If Text1 = Empty Then
Text1.SetFocus
Exit Sub
End If
str1 = Text1.Text
    strsearch = "pin like '" & str1 & "'"

    Adodc1.Recordset.MoveFirst
    Adodc1.Recordset.Filter = (strsearch)
    If Text1.Text = Text3.Text Then
'MsgBox "welcome " & Adodc1.Recordset.Fields(0)
Label5.Caption = Adodc1.Recordset.Fields(0)
Label3.Caption = Adodc1.Recordset.Fields(3)
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command15.Enabled = True
Label63.Visible = True
Label3.Visible = True
Label5.Visible = True
MediaPlayer1.FileName = "C:\Microsoft Visual Studio\VB98\bank project\sounds\buttons\button7.wav"
MediaPlayer1.Play
Else
MediaPlayer1.FileName = "C:\Microsoft Visual Studio\VB98\bank project\sounds\buttons\button11.wav"
MediaPlayer1.Play
MsgBox "pin is not valid,transaction canceled"
End
End If
Text1.Text = Empty
Text1.SetFocus
End Sub

Private Sub Command10_Click()
If Label17.Caption <> none Then  'And Label17.Caption <> Empty Then'
Adodc1.Recordset.Fields(5) = "none"
Adodc1.Recordset.Fields(3) = Val(Adodc1.Recordset.Fields(3)) + Val(Adodc1.Recordset.Fields(7))
Adodc1.Recordset.Fields(3) = Adodc1.Recordset.Fields(3) - Adodc1.Recordset.Fields(7) * 0.05
Adodc1.Recordset.Fields(7) = Empty
Combo2.Visible = False
Text11.Visible = False
Text11.Text = "0"
End If
Adodc1.Recordset.Save
End Sub

Private Sub Command11_Click()
If Combo2 = Empty Then
Combo2.Visible = True
Label40.Visible = True
Label41.Visible = True
Text11.Visible = True
MsgBox "please choose the package type and the amount to deposit for that package"
End If
If Combo2 = "1 year pack" And Text11.Text <> Empty Then
Adodc1.Recordset.Fields(5) = "1 year pack"
'Adodc1.Recordset.Save
Combo2.Visible = False
Label40.Visible = False
Label41.Visible = False
Text11.Visible = False
End If
If Combo2 = "3 year pack" And Text11.Text <> Empty Then
Adodc1.Recordset.Fields(5) = "3 year pack"
'Adodc1.Recordset.Save
Combo2.Visible = False
Label40.Visible = False
Label41.Visible = False
Text11.Visible = False
End If
If Combo2 = "5 year pack" And Text11.Text <> Empty Then
Adodc1.Recordset.Fields(5) = "5 year pack"
'Adodc1.Recordset.Save
Combo2.Visible = False
Label40.Visible = False
Label41.Visible = False
Text11.Visible = False
End If
If Combo2 = "10 year pack" And Text11.Text <> Empty Then
Adodc1.Recordset.Fields(5) = "10 year pack"
'Adodc1.Recordset.Save
Combo2.Visible = False
Label40.Visible = False
Label41.Visible = False
Text11.Visible = False
End If
If Val(Text11.Text) > Val(Label52.Caption) Then
MediaPlayer1.FileName = "C:\Microsoft Visual Studio\VB98\bank project\sounds\buttons\combine_button_locked.wav"
MediaPlayer1.Play
MsgBox "not enough funds, transaction canceled"
Adodc1.Recordset.Fields(5) = "none"
Combo2 = Empty
Else
Label3.Caption = Val(Label52.Caption) - Val(Text11.Text)
Adodc1.Recordset.Fields(3) = Val(Adodc1.Recordset.Fields(3)) - Val(Text11.Text)
Adodc1.Recordset.Fields(7) = Text11.Text
Combo2 = Empty
End If
Adodc1.Recordset.Save
End Sub

Private Sub Command12_Click()
Adodc1.Recordset.Save
End Sub

Private Sub Command13_Click()
MsgBox "at any time you deactivate the packages (except the default one),your package subscription is canceled and you get no extra balance BUT you will be fined of 5% of your balalnce "
End Sub

Private Sub Command14_Click()
SSTab1.Visible = False
Frame1.Visible = True
Label14.Visible = False
End Sub

Private Sub Command15_Click()
Adodc1.Recordset.Fields(3) = Val(Adodc1.Recordset.Fields(3)) + Val(Text8.Text)
'Text5.Text = Val(Text5.Text) + Val(Text8.Text)
Dim variable1 As String
Dim a As String
Dim b As String
b = Adodc1.Recordset.Fields(2)
a = Adodc1.Recordset.Fields(9)
Open "C:\Microsoft Visual Studio\VB98\bank project\deposit\" & a & b & ".Txt" For Output As #1
Print #1, Text8
Close #1
Adodc1.Recordset.Fields(9) = Val(Adodc1.Recordset.Fields(9)) + 1
Adodc1.Recordset.Save
MediaPlayer1.FileName = "C:\Microsoft Visual Studio\VB98\bank project\sounds\buttons\coinslot1.wav"
MediaPlayer1.Play
End Sub

Private Sub Command16_Click()
Adodc1.Recordset.Save
End
End Sub

Private Sub Command17_Click()
Do While List2.ListCount <> 0
     List2.RemoveItem (0)
    Loop
   Do While List1.ListCount <> 0
    List1.RemoveItem (0)
    Loop
    Timer10.Enabled = True
    Timer11.Enabled = True
End Sub

Private Sub Command18_Click()
Unload Form1
Load Form1
Form1.Show
End Sub

Private Sub Command19_Click()
Dim usermsg As String
Dim usermsg2 As String
usermsg = InputBox("input pin", "pin code", "Enter your pin here", 500, 700)
If usermsg = Adodc1.Recordset.Fields(1) Then
usermsg2 = InputBox("input new pin", "pin code", "Enter your new pin here", 500, 700)
Adodc1.Recordset.Fields(1) = usermsg2
End If
End Sub

Private Sub Command2_Click()
Dim usermsg As String
Dim usermsg2 As String
usermsg = InputBox("input pin", "pin code", "Enter your pin here", 500, 700)
If usermsg = Adodc1.Recordset.Fields(1) Then
usermsg2 = InputBox("input new pin", "pin code", "Enter your new pin here", 500, 700)
Adodc1.Recordset.Fields(1) = usermsg2
End If
End Sub


Private Sub Command20_Click()
On Error GoTo err
Adodc1.Recordset.Save
Exit Sub
err:
MsgBox "failed to save"
End Sub

Private Sub Command3_Click()
If Val(Text7.Text) > Val(Label3.Caption) Then
MediaPlayer1.FileName = "C:\Microsoft Visual Studio\VB98\bank project\sounds\buttons\combine_button_locked.wav"
MediaPlayer1.Play
MsgBox "not enough funds, transaction canceled"

Else
Dim variable1 As String
Dim a As String
Dim b As String
b = Adodc1.Recordset.Fields(2)
a = Adodc1.Recordset.Fields(8)
Open "C:\Microsoft Visual Studio\VB98\bank project\withdraw\" & a & b & ".Txt" For Output As #1
Print #1, Text7
Close #1
MediaPlayer1.FileName = "C:\Microsoft Visual Studio\VB98\bank project\sounds\buttons\button4.wav"
MediaPlayer1.Play
Adodc1.Recordset.Fields(8) = Val(Adodc1.Recordset.Fields(8)) + 1
Label3.Caption = Val(Label3.Caption) - Val(Text7.Text)
Adodc1.Recordset.Fields(3) = Val(Adodc1.Recordset.Fields(3)) - Val(Text7.Text)
End If
Adodc1.Recordset.Save
End Sub

Private Sub Command4_Click()
Label5.Visible = False
Label3.Visible = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command15.Enabled = False
Label63.Visible = False
MediaPlayer1.FileName = "C:\Microsoft Visual Studio\VB98\bank project\sounds\buttons\button6.wav"
MediaPlayer1.Play
'Adodc1.Recordset.Save
Command5.Enabled = True
End Sub


Private Sub Command5_Click()
Frame1.Visible = False
Text9.SetFocus
If Text9.Text = Empty Then
Exit Sub
End If
str1 = Text9.Text
    strsearch = "pin like '" & str1 & "'"

    Adodc1.Recordset.MoveFirst
    Adodc1.Recordset.Filter = (strsearch)
    If Text9.Text = Text3.Text Then
    SSTab1.Visible = True

    Label13.Caption = Adodc1.Recordset.Fields(0) & "  is online"

    Label13.Visible = True
    Command5.Caption = "change account"
    Do While List2.ListCount <> 0
     List2.RemoveItem (0)
    Loop
   Do While List1.ListCount <> 0
    List1.RemoveItem (0)
    Loop
    Timer10.Enabled = True
    Timer11.Enabled = True
    Label70.Caption = Adodc1.Recordset.Fields(3)
    
   MediaPlayer1.FileName = "C:\Microsoft Visual Studio\VB98\bank project\sounds\buttons\button3.wav"
MediaPlayer1.Play
    Else
    MediaPlayer1.FileName = "C:\Microsoft Visual Studio\VB98\bank project\sounds\buttons\button11.wav"
MediaPlayer1.Play
    MsgBox "pin is not valid.transaction cancelled"
    
End
    End If
    Timer1.Enabled = True
    Timer2.Enabled = True
    Timer3.Enabled = True
    Timer4.Enabled = True
    Timer5.Enabled = True
    Timer6.Enabled = True
Text9.SetFocus
Text9.Text = Empty
End Sub







Private Sub Command6_Click()
Label70.Caption = Adodc1.Recordset.Fields(3)
Label21.Caption = 0
Label33.Caption = 0
End Sub

Private Sub Command7_Click()
Label25.Caption = Val(Label25.Caption) + 6
Label21.Caption = Val(Label21.Caption) + 1
End Sub

Private Sub Command8_Click()
Adodc1.Recordset.Fields(3) = Adodc1.Recordset.Fields(3) + Val(Label18.Caption) + Val(Label60.Caption)
MediaPlayer1.FileName = "C:\Microsoft Visual Studio\VB98\bank project\sounds\buttons\coinslot1.wav"
MediaPlayer1.Play
Adodc1.Recordset.Save

End Sub

Private Sub Command9_Click()
Timer3.Enabled = True
End Sub

Private Sub Form_Load()
Label24.Caption = Val(Label24.Caption) + 16


End Sub

Private Sub Label14_Change()
On Error GoTo err:
Adodc1.Recordset.Fields(3) = Int(Adodc1.Recordset.Fields(3))
Exit Sub
err:
End Sub

Private Sub Label21_Change()
Label33.Caption = Val(Label21.Caption) / 2
End Sub

Private Sub Label21_Click()
'1 session = 6 months'
End Sub

Private Sub Label24_Change()
If Val(Label24.Caption) > 30 Then
Label24.Caption = Val(Label24.Caption) - 30
Label25.Caption = Val(Label25.Caption) + 1
End If
End Sub

Private Sub Label25_Change()
If Val(Label25.Caption) >= 6 Then
X = Label70.Caption
Label70.Caption = Val(Label70.Caption) + Val(Label18.Caption) + Val(Label60.Caption)
X = X + Val(Label18.Caption) + Val(Label60.Caption)
Label25.Caption = Empty
MsgBox "your predicted balance is!" & X
End If
End Sub

Private Sub Option1_Click()

End Sub

Private Sub Label69_Change()
'On Error GoTo err
'Adodc1.Recordset.Save
'Exit Sub
'err:
'MsgBox "label69(amount) failed to save"
End Sub

Private Sub Text16_Change()
'On Error GoTo err
'Adodc1.Recordset.Save
'Exit Sub
'err:
'MsgBox "text16(name) failed to save"
End Sub

Private Sub Text19_Change()
'On Error GoTo err
'Adodc1.Recordset.Save
'Exit Sub
'err:
'MsgBox "text19(address) failed to save"
End Sub

Private Sub Text20_Change()

End Sub

Private Sub Text21_Change()
'On Error GoTo err
'Adodc1.Recordset.Save
'Exit Sub
'err:
'MsgBox "text21(acc_ref) failed to save"
End Sub

Private Sub Timer1_Timer()
Label18.Caption = Val(Label14.Caption) * Val(Label16.Caption) / 200
End Sub

Private Sub Timer10_Timer()
On Error GoTo err:
X = Adodc1.Recordset.Fields(0)
Y = 0
Do While Y <= X
Y = Y + 1
Dim variable2 As String
Dim a As Integer
Dim b As String
For a = 0 To 10000
b = Adodc1.Recordset.Fields(2)
Open "C:\Microsoft Visual Studio\VB98\bank project\deposit\" & a & b & ".Txt" For Input As #1
Input #1, variable2
List2.AddItem (variable2)
Close #1
Next a
Loop
Exit Sub
err:
Timer10.Enabled = False
End Sub

Private Sub Timer11_Timer()
On Error GoTo err:
X = Form1.Adodc1.Recordset.Fields(0)
Y = 0
Do While Y <= X
Y = Y + 1
Dim variable1 As String
Dim a As Integer
Dim b As String
For a = 0 To 10000
b = Form1.Adodc1.Recordset.Fields(2)
Open "C:\Microsoft Visual Studio\VB98\bank project\withdraw\" & a & b & ".Txt" For Input As #1
Input #1, variable1
List1.AddItem (variable1)
Close #1
Next a
Loop
Exit Sub
err:
Timer11.Enabled = False
End Sub

Private Sub Timer2_Timer()
'Label14.Caption = Text5.Text
'Label16.Caption = Text6.Text
'Label17.Caption = Combo1
'Label52.Caption = Text5.Text
'If Label17.Caption = Empty Then
'Label15.Caption = "6 months(default)"
End Sub

Private Sub Timer3_Timer()
If Val(Label70.Caption) < Val(Text10.Text) Then
Label25.Caption = Val(Label25.Caption) + 6
Label21.Caption = Val(Label21.Caption) + 1
Else
Timer3.Enabled = False
End If
End Sub

Private Sub Timer4_Timer()
'If Combo1 = "1 year pack" Then Text12.Text = "13"
'If Combo1 = "3 year pack" Then Text12.Text = "22"

'If Combo1 = "5 year pack" Then Text12.Text = "30"

'If Combo1 = "10 year pack" Then Text12.Text = "50"

'If Combo1 = "none" Then Text12.Text = ""
End Sub

Private Sub Timer5_Timer()
Label60.Caption = Val(Label59.Caption) * Val(Label57.Caption) / 200
End Sub

Private Sub Timer6_Timer()
'Label57.Caption = Text12.Text
'Label59.Caption = Text13.Text
'Label54.Caption = Combo1
End Sub


Private Sub Timer8_Timer()
Timer8.Enabled = False
End Sub

Private Sub Timer9_Timer()
On Error GoTo err:
If Adodc1.Recordset.Fields(5) = "none" Then
Command11.Enabled = True
Command10.Enabled = False
Else
Command11.Enabled = False
Command10.Enabled = True
End If
Exit Sub
err:
End Sub
