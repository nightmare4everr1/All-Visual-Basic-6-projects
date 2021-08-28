VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9930
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   ScaleHeight     =   9930
   ScaleWidth      =   11400
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command18 
      Caption         =   "refresh"
      Height          =   495
      Left            =   8400
      TabIndex        =   93
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Command17"
      Height          =   495
      Left            =   7320
      TabIndex        =   92
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton Command16 
      Caption         =   "exit"
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
      Left            =   120
      TabIndex        =   91
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton Command14 
      Caption         =   "change account"
      Height          =   495
      Left            =   2880
      TabIndex        =   89
      Top             =   1800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text13 
      DataField       =   "cashinpack"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1440
      TabIndex        =   80
      Text            =   "Text13"
      Top             =   8880
      Width           =   1215
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1440
      Top             =   9240
   End
   Begin VB.TextBox Text12 
      DataField       =   "packinterest"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1440
      TabIndex        =   78
      Text            =   "Text12"
      Top             =   8520
      Width           =   1215
   End
   Begin VB.CommandButton Command12 
      Caption         =   "save"
      Height          =   315
      Left            =   120
      TabIndex        =   54
      Top             =   9600
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7335
      Left            =   3000
      TabIndex        =   28
      Top             =   2400
      Visible         =   0   'False
      Width           =   8565
      _ExtentX        =   15108
      _ExtentY        =   12938
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "general details"
      TabPicture(0)   =   "Form1.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label33"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label32"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label31"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label21"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label20"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label30"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label29"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label28"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label27"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label26"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label25"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label24"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label23"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label22"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label19"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label18"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label17"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label16"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label15"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label14"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Line1"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label53"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label54"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label55"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label56"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label57"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Label58"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Label59"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Label60"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Timer3"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Command9"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Text10"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Command7"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Timer2"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Command8"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Timer1"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Command6"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Timer5"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Timer6"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).ControlCount=   39
      TabCaption(1)   =   "package details"
      TabPicture(1)   =   "Form1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text11"
      Tab(1).Control(1)=   "Command10"
      Tab(1).Control(2)=   "Command11"
      Tab(1).Control(3)=   "Combo2"
      Tab(1).Control(4)=   "Command13"
      Tab(1).Control(5)=   "Label52"
      Tab(1).Control(6)=   "Label51"
      Tab(1).Control(7)=   "Label34"
      Tab(1).Control(8)=   "Label35"
      Tab(1).Control(9)=   "Label36"
      Tab(1).Control(10)=   "Label37"
      Tab(1).Control(11)=   "Label38"
      Tab(1).Control(12)=   "Label39"
      Tab(1).Control(13)=   "Label40"
      Tab(1).Control(14)=   "Label41"
      Tab(1).Control(15)=   "Label44"
      Tab(1).Control(16)=   "Label45"
      Tab(1).Control(17)=   "Label46"
      Tab(1).Control(18)=   "Label47"
      Tab(1).Control(19)=   "Label48"
      Tab(1).Control(20)=   "Label49"
      Tab(1).Control(21)=   "Label50"
      Tab(1).ControlCount=   22
      TabCaption(2)   =   "extras"
      TabPicture(2)   =   "Form1.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.Timer Timer6 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   4560
         Top             =   5880
      End
      Begin VB.Timer Timer5 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   5880
         Top             =   5160
      End
      Begin VB.TextBox Text11 
         Height          =   525
         Left            =   -68520
         TabIndex        =   74
         Top             =   1320
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton Command10 
         Caption         =   "deactivate package"
         Height          =   495
         Left            =   -71760
         TabIndex        =   58
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CommandButton Command11 
         Caption         =   "activate package"
         Height          =   495
         Left            =   -71760
         TabIndex        =   57
         Top             =   600
         Width           =   1575
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "Form1.frx":0054
         Left            =   -68640
         List            =   "Form1.frx":0064
         TabIndex        =   56
         Top             =   600
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton Command13 
         Caption         =   "click to see terms and conditions"
         Height          =   615
         Left            =   -69960
         TabIndex        =   55
         Top             =   3960
         Width           =   1335
      End
      Begin VB.CommandButton Command6 
         Caption         =   "refresh data"
         Height          =   495
         Left            =   3600
         TabIndex        =   53
         Top             =   3840
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   4440
         Top             =   2040
      End
      Begin VB.CommandButton Command8 
         Caption         =   "time travel"
         Height          =   495
         Left            =   3600
         TabIndex        =   32
         Top             =   4320
         Width           =   975
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   4920
         Top             =   2040
      End
      Begin VB.CommandButton Command7 
         Caption         =   "predict balance after next session"
         Height          =   495
         Left            =   120
         TabIndex        =   31
         Top             =   2760
         Width           =   1695
      End
      Begin VB.TextBox Text10 
         Height          =   495
         Left            =   2400
         TabIndex        =   30
         Top             =   3360
         Width           =   1095
      End
      Begin VB.CommandButton Command9 
         Caption         =   "predict sessions needed"
         Height          =   495
         Left            =   3600
         TabIndex        =   29
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   5400
         Top             =   2040
      End
      Begin VB.Label Label60 
         Caption         =   "Label60"
         Height          =   375
         Left            =   4920
         TabIndex        =   88
         Top             =   5280
         Width           =   855
      End
      Begin VB.Label Label59 
         Caption         =   "Label59"
         Height          =   375
         Left            =   3600
         TabIndex        =   87
         Top             =   5640
         Width           =   1335
      End
      Begin VB.Label Label58 
         Caption         =   "amount deposited"
         Height          =   375
         Left            =   2280
         TabIndex        =   86
         Top             =   5640
         Width           =   1335
      End
      Begin VB.Label Label57 
         Caption         =   "Label57"
         Height          =   375
         Left            =   1320
         TabIndex        =   85
         Top             =   5640
         Width           =   615
      End
      Begin VB.Label Label56 
         Caption         =   "interest rate"
         Height          =   375
         Left            =   120
         TabIndex        =   84
         Top             =   5640
         Width           =   1215
      End
      Begin VB.Label Label55 
         Caption         =   "after one session ="
         Height          =   255
         Left            =   3480
         TabIndex        =   83
         Top             =   5280
         Width           =   1455
      End
      Begin VB.Label Label54 
         Caption         =   "Label54"
         Height          =   255
         Left            =   2280
         TabIndex        =   82
         Top             =   5280
         Width           =   1335
      End
      Begin VB.Label Label53 
         Caption         =   "predicted cash of package "
         Height          =   375
         Left            =   120
         TabIndex        =   81
         Top             =   5280
         Width           =   2295
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   5880
         Y1              =   5040
         Y2              =   5040
      End
      Begin VB.Label Label52 
         Caption         =   "Label52"
         DataField       =   "amount"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   -68640
         TabIndex        =   76
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label51 
         Caption         =   "current balance"
         Height          =   375
         Left            =   -69960
         TabIndex        =   75
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         Caption         =   "3 year pack"
         Height          =   495
         Left            =   -74760
         TabIndex        =   73
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         Caption         =   "the conditions are the same as of the previous pack,but you get 22%extra surplus"
         Height          =   495
         Left            =   -73320
         TabIndex        =   72
         Top             =   3960
         Width           =   3135
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         Caption         =   "5 year pack"
         Height          =   495
         Left            =   -74760
         TabIndex        =   71
         Top             =   4680
         Width           =   1215
      End
      Begin VB.Label Label37 
         Alignment       =   2  'Center
         Caption         =   "same except you get 30% extra"
         Height          =   495
         Left            =   -73320
         TabIndex        =   70
         Top             =   4680
         Width           =   3135
      End
      Begin VB.Label Label38 
         Alignment       =   2  'Center
         Caption         =   "ten year pack"
         Height          =   495
         Left            =   -74760
         TabIndex        =   69
         Top             =   5520
         Width           =   1215
      End
      Begin VB.Label Label39 
         Alignment       =   2  'Center
         Caption         =   "same except you get 50% extra"
         Height          =   495
         Left            =   -73200
         TabIndex        =   68
         Top             =   5520
         Width           =   3135
      End
      Begin VB.Label Label40 
         Caption         =   "deposit amount"
         Height          =   255
         Left            =   -69960
         TabIndex        =   67
         Top             =   1440
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label41 
         Caption         =   "package type"
         Height          =   495
         Left            =   -69960
         TabIndex        =   66
         Top             =   600
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label44 
         Alignment       =   2  'Center
         Caption         =   "your current package is"
         Height          =   495
         Left            =   -74760
         TabIndex        =   65
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label45 
         Alignment       =   2  'Center
         Caption         =   "Label45"
         DataField       =   "packages"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   -73320
         TabIndex        =   64
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label46 
         Alignment       =   2  'Center
         Caption         =   "package types"
         Height          =   495
         Left            =   -74760
         TabIndex        =   63
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label47 
         Alignment       =   2  'Center
         Caption         =   "none(default)"
         Height          =   495
         Left            =   -74760
         TabIndex        =   62
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label48 
         Alignment       =   2  'Center
         Caption         =   "this is the default package.at the end of every session you get 10% surplus cash of your deposited amount"
         Height          =   495
         Left            =   -73320
         TabIndex        =   61
         Top             =   2040
         Width           =   3855
      End
      Begin VB.Label Label49 
         Alignment       =   2  'Center
         Caption         =   "1 year pack"
         Height          =   495
         Left            =   -74760
         TabIndex        =   60
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label50 
         Alignment       =   2  'Center
         Caption         =   $"Form1.frx":009D
         Height          =   855
         Left            =   -73440
         TabIndex        =   59
         Top             =   2760
         Width           =   5295
      End
      Begin VB.Label Label14 
         Height          =   375
         Left            =   1440
         TabIndex        =   52
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label15 
         Height          =   735
         Left            =   4680
         TabIndex        =   51
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label16 
         Height          =   615
         Left            =   4680
         TabIndex        =   50
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label17 
         Height          =   735
         Left            =   1440
         TabIndex        =   49
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label18 
         Height          =   735
         Left            =   2400
         TabIndex        =   48
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label Label19 
         Caption         =   "time passed since cash deposited"
         Height          =   495
         Left            =   240
         TabIndex        =   47
         Top             =   4200
         Width           =   1695
      End
      Begin VB.Label Label22 
         Caption         =   "day number"
         Height          =   255
         Left            =   2160
         TabIndex        =   46
         Top             =   4200
         Width           =   975
      End
      Begin VB.Label Label23 
         Caption         =   "months"
         Height          =   375
         Left            =   2160
         TabIndex        =   45
         Top             =   4680
         Width           =   1095
      End
      Begin VB.Label Label24 
         Caption         =   "25"
         Height          =   255
         Left            =   3240
         TabIndex        =   44
         Top             =   4200
         Width           =   255
      End
      Begin VB.Label Label25 
         Caption         =   "0"
         Height          =   255
         Left            =   3240
         TabIndex        =   43
         Top             =   4680
         Width           =   255
      End
      Begin VB.Label Label26 
         Caption         =   "period of deposit"
         Height          =   495
         Left            =   3360
         TabIndex        =   42
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label27 
         Caption         =   "interest rate"
         Height          =   495
         Left            =   3360
         TabIndex        =   41
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label28 
         Caption         =   "your current account"
         Height          =   495
         Left            =   120
         TabIndex        =   40
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label29 
         Caption         =   "package"
         Height          =   615
         Left            =   120
         TabIndex        =   39
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label30 
         Caption         =   "predictid surplus after 6 months(one session)"
         Height          =   615
         Left            =   120
         TabIndex        =   38
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label Label20 
         Caption         =   "session number"
         Height          =   255
         Left            =   1920
         TabIndex        =   37
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label21 
         Height          =   375
         Left            =   3120
         TabIndex        =   36
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label31 
         Caption         =   "predict how many sessions are needed to achive the target of"
         Height          =   495
         Left            =   120
         TabIndex        =   35
         Top             =   3360
         Width           =   2295
      End
      Begin VB.Label Label32 
         Caption         =   "years needed"
         Height          =   375
         Left            =   4080
         TabIndex        =   34
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label33 
         Height          =   375
         Left            =   5160
         TabIndex        =   33
         Top             =   2880
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   4935
      Left            =   0
      TabIndex        =   16
      Top             =   1320
      Width           =   2655
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   1560
         TabIndex        =   95
         Text            =   "Text8"
         Top             =   3960
         Width           =   975
      End
      Begin VB.CommandButton Command15 
         Caption         =   "deposit aditional amount"
         Enabled         =   0   'False
         Height          =   495
         Left            =   120
         TabIndex        =   94
         Top             =   3840
         Width           =   1335
      End
      Begin VB.Timer Timer7 
         Interval        =   1
         Left            =   2040
         Top             =   1440
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
         Caption         =   "enter pin"
         Height          =   495
         Left            =   720
         TabIndex        =   21
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "change pin code"
         Enabled         =   0   'False
         Height          =   495
         Left            =   120
         TabIndex        =   20
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "withdraw cash ammount "
         Enabled         =   0   'False
         Height          =   495
         Left            =   1320
         TabIndex        =   19
         Top             =   2880
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "end transaction"
         Enabled         =   0   'False
         Height          =   255
         Left            =   720
         TabIndex        =   18
         Top             =   4560
         Width           =   1455
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   1320
         TabIndex        =   17
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   2640
         Y1              =   4440
         Y2              =   4440
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   2640
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Label Label2 
         Caption         =   "your current balance"
         Height          =   495
         Left            =   120
         TabIndex        =   27
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   960
         TabIndex        =   26
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "account"
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   840
         TabIndex        =   24
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label12 
         Caption         =   "credit card machine"
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
      Left            =   5520
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
      ItemData        =   "Form1.frx":01AC
      Left            =   1440
      List            =   "Form1.frx":01BF
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   8160
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      DataField       =   "interest"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Text            =   "Text6"
      Top             =   7800
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      DataField       =   "amount"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Text            =   "Text5"
      Top             =   7440
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      DataField       =   "roll"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Text            =   "Text4"
      Top             =   6960
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      DataField       =   "pin"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   6600
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      DataField       =   "name"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   6360
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   120
      Top             =   9240
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
   Begin VB.Label Label61 
      BackStyle       =   0  'Transparent
      Caption         =   "is online"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   90
      Top             =   1080
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Label43 
      Caption         =   "deposited amount in pack"
      Height          =   375
      Left            =   120
      TabIndex        =   79
      Top             =   8760
      Width           =   1215
   End
   Begin VB.Label Label42 
      Caption         =   "package interest"
      Height          =   375
      Left            =   120
      TabIndex        =   77
      Top             =   8520
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
      ForeColor       =   &H8000000D&
      Height          =   855
      Left            =   3240
      TabIndex        =   14
      Top             =   960
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label Label11 
      Caption         =   "package"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   8160
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "interest"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   7800
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "ammount"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "roll"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "pin"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "name"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "askari bank"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3720
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
If Text1 = Empty Then
Exit Sub
End If
str1 = Text1.Text
    strsearch = "pin like '" & str1 & "'"

    Adodc1.Recordset.MoveFirst
    Adodc1.Recordset.Filter = (strsearch)
    If Text1.Text = Text3.Text Then
MsgBox "correct"
Label5.Caption = Text2.Text
Label3.Caption = Text5.Text
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command15.Enabled = True
Timer7.Enabled = True
Else
MsgBox "pin is not valid,transaction canceled"
End
End If
End Sub

Private Sub Command10_Click()
If Label17.Caption <> none And Label17.Caption <> Empty Then
Combo1 = "none"
Text5.Text = Val(Text5.Text) + Val(Text13.Text)
Text13.Text = Empty
End If
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
Combo1 = "1 year pack"
Else
If Combo2 = "3 year pack" And Text11.Text <> Empty Then
Combo1 = "3 year pack"
Else
If Combo2 = "5 year pack" And Text11.Text <> Empty Then
Combo1 = "5 year pack"
Else
If Combo2 = "10 year pack" And Text11.Text <> Empty Then
Combo1 = "10 year pack"
End If
End If
End If
End If
If Val(Text11.Text) > Val(Label52.Caption) Then
MsgBox "not enough funds, transaction canceled"
Combo1 = Empty
Combo2 = Empty
Else
Label3.Caption = Val(Label52.Caption) - Val(Text11.Text)
Text5.Text = Val(Text5.Text) - Val(Text11.Text)
Text13.Text = Text11.Text
End If
End Sub

Private Sub Command12_Click()
Adodc1.Recordset.Save
End Sub

Private Sub Command13_Click()
MsgBox "at any time you deactivate the packages (except the default one),your package subscription is canceled and you get no extra balance "
End Sub

Private Sub Command14_Click()
MsgBox "run program again and click on acces your account"
End
End Sub

Private Sub Command15_Click()
Text5.Text = Val(Text5.Text) + Val(Text8.Text)
End Sub

Private Sub Command16_Click()
Adodc1.Recordset.Save
End
End Sub

Private Sub Command17_Click()
Timer1.Enabled = True
    Timer2.Enabled = False
    Timer3.Enabled = False
    Timer4.Enabled = False
    Timer5.Enabled = False
    Timer6.Enabled = False
    Timer7.Enabled = False
End Sub

Private Sub Command18_Click()
Unload Form1
Load Form1
Form1.Show
End Sub

Private Sub Command2_Click()
Dim usermsg As String
Dim usermsg2 As String
usermsg = InputBox("input pin", "pin code", "Enter your pin here", 500, 700)
If usermsg = Text3.Text Then
usermsg2 = InputBox("input new pin", "pin code", "Enter your new pin here", 500, 700)
Text3.Text = usermsg2
End If
End Sub


Private Sub Command3_Click()
If Val(Text7.Text) > Val(Label3.Caption) Then
MsgBox "not enough funds, transaction canceled"
Else
Label3.Caption = Val(Label3.Caption) - Val(Text7.Text)
Text5.Text = Val(Text5.Text) - Val(Text7.Text)
End If
End Sub

Private Sub Command4_Click()
Timer7.Enabled = False
Label5.Caption = Empty
Label3.Caption = Empty
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command15.Enabled = False
Text1.Text = Empty
Adodc1.Recordset.Save
End Sub


Private Sub Command5_Click()
If Text9.Text = Empty Then
Exit Sub
End If
str1 = Text9.Text
    strsearch = "pin like '" & str1 & "'"

    Adodc1.Recordset.MoveFirst
    Adodc1.Recordset.Filter = (strsearch)
    If Text9.Text = Text3.Text Then
    SSTab1.Visible = True
    Command5.Enabled = False
    Command14.Visible = True
    Label13.Caption = Text2.Text
    Label61.Visible = True
    Label13.Visible = True
    Else
    MsgBox "pin is not valid.transaction cancelled"
    End
    End If
    Timer1.Enabled = True
    Timer2.Enabled = True
    Timer3.Enabled = True
    Timer4.Enabled = True
    Timer5.Enabled = True
    Timer6.Enabled = True
    Timer7.Enabled = True
End Sub

Private Sub Command6_Click()
Label14.Caption = Text5.Text
Label16.Caption = Text6.Text
Label17.Caption = Combo1
If Label17.Caption = Empty Then
Label15.Caption = "6 months(default)"
End If
End Sub

Private Sub Command7_Click()
Label25.Caption = Val(Label25.Caption) + 6
Label21.Caption = Val(Label21.Caption) + 1
End Sub

Private Sub Command8_Click()
Label24.Caption = Val(Label24.Caption) + 16
End Sub

Private Sub Command9_Click()
Timer3.Enabled = True
End Sub

Private Sub Form_Load()
Label24.Caption = Val(Label24.Caption) + 16
End Sub

Private Sub Label21_Change()
Label33.Caption = Val(Label21.Caption) / 2
End Sub

Private Sub Label21_Click()
'1 session = 6 months'
End Sub

Private Sub Label24_Change()
If Val(Label24.Caption) > 30 Then
Label25.Caption = Val(Label25.Caption) + 1
Label24.Caption = Val(Label24.Caption) - 30
End If
End Sub

Private Sub Label25_Change()
If Val(Label25.Caption) >= 6 Then
Label14.Caption = Val(Label14.Caption) + Val(Label18.Caption) + Val(Label60.Caption)
Text5.Text = Val(Text5.Text) + Val(Label18.Caption) + Val(Label60.Caption)
Label25.Caption = Empty
MsgBox "you have recieved the interest and is now added to your balance!"
End If
End Sub

Private Sub Option1_Click()

End Sub

Private Sub Timer1_Timer()
Label18.Caption = Val(Label14.Caption) * Val(Label16.Caption) / 200
End Sub

Private Sub Timer2_Timer()
Label14.Caption = Text5.Text
Label16.Caption = Text6.Text
Label17.Caption = Combo1
Label52.Caption = Text5.Text
If Label17.Caption = Empty Then
Label15.Caption = "6 months(default)"
End If
End Sub

Private Sub Timer3_Timer()
If Val(Label14.Caption) < Val(Text10.Text) Then
Label25.Caption = Val(Label25.Caption) + 6
Label21.Caption = Val(Label21.Caption) + 1
Else
Timer3.Enabled = False
End If
End Sub

Private Sub Timer4_Timer()
If Combo1 = "1 year pack" Then Text12.Text = "13"
If Combo1 = "3 year pack" Then Text12.Text = "22"

If Combo1 = "5 year pack" Then Text12.Text = "30"

If Combo1 = "10 year pack" Then Text12.Text = "50"

If Combo1 = "none" Then Text12.Text = ""
End Sub

Private Sub Timer5_Timer()
Label60.Caption = Val(Label59.Caption) * Val(Label57.Caption) / 200
End Sub

Private Sub Timer6_Timer()
Label57.Caption = Text12.Text
Label59.Caption = Text13.Text
Label54.Caption = Combo1
End Sub

Private Sub Timer7_Timer()
Label3 = Text5
End Sub

Private Sub Timer8_Timer()
If SSTab1.Visible = True Then
Timer1.Enabled = True
    Timer2.Enabled = True
    Timer3.Enabled = True
    Timer4.Enabled = True
    Timer5.Enabled = True
    Timer6.Enabled = True
    Timer7.Enabled = True
    End If
End Sub
