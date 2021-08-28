VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{A40B4820-D5FD-11D1-8818-C199198E9702}#1.8#0"; "MMTOOLSX.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "traffic flow"
   ClientHeight    =   11160
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   15270
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Form1.frx":0000
   MousePointer    =   99  'Custom
   Moveable        =   0   'False
   ScaleHeight     =   11160
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MMToolsX.MMSliderX MMSliderX2 
      Height          =   1095
      Left            =   14280
      TabIndex        =   67
      Top             =   8280
      Width           =   255
      Object.Width           =   17
      Object.Height          =   73
      Color           =   -2147483633
      Enabled         =   0
      Bevel.BorderWidth=   5
      MaxValue        =   3
      Orientation     =   0
      ThumbWidth      =   23
      ThumbHeight     =   11
      Scale.Visible   =   0
      Scale.Color     =   0
      Scale.Style     =   0
      Scale.TickCount =   11
      Scale.EnlargeEvery=   5
      Scale.Size      =   7
      Scale.Origin    =   1
      Scale.Connect   =   1
      _Handle         =   7
   End
   Begin MMToolsX.MMWheelX MMWheelX2 
      Height          =   900
      Left            =   14280
      TabIndex        =   56
      Top             =   7320
      Width           =   900
      Object.Width           =   60
      Object.Height          =   60
      Bevel.BevelOuter=   0
      Color           =   -2147483633
      AutoSize        =   1
      BackBmp.Data    =   $"Form1.frx":030A
      MinValue        =   1
      Value           =   10
      Scale.Visible   =   1
      Scale.Color     =   0
      Scale.Style     =   0
      Scale.TickCount =   11
      Scale.EnlargeEvery=   5
      Scale.Size      =   7
      Scale.Origin    =   0
      Scale.Connect   =   0
      Radius          =   12
      _Handle         =   5
   End
   Begin MMToolsX.MMWheelX MMWheelX1 
      Height          =   900
      Left            =   13320
      TabIndex        =   55
      Top             =   7320
      Width           =   900
      Object.Width           =   60
      Object.Height          =   60
      Bevel.BevelOuter=   0
      Color           =   -2147483637
      AutoSize        =   1
      BackBmp.Data    =   $"Form1.frx":18B9
      MinValue        =   1
      Value           =   10
      Scale.Visible   =   1
      Scale.Color     =   0
      Scale.Style     =   0
      Scale.TickCount =   11
      Scale.EnlargeEvery=   5
      Scale.Size      =   7
      Scale.Origin    =   0
      Scale.Connect   =   0
      Radius          =   12
      _Handle         =   4
   End
   Begin MMToolsX.MMGaugeX MMGaugeX2 
      Height          =   495
      Left            =   3240
      TabIndex        =   40
      Top             =   10440
      Width           =   5895
      ForeColor       =   8454016
      Font.Charset    =   0
      Font.Color      =   49152
      Font.Height     =   -29
      Font.Name       =   "Comic Sans MS"
      Font.Style      =   3
      ParentFont      =   0
      _Handle         =   2
   End
   Begin MMToolsX.MMGaugeX MMGaugeX1 
      Height          =   495
      Left            =   3240
      TabIndex        =   38
      Top             =   9840
      Width           =   5895
      ForeColor       =   -2147483635
      Font.Charset    =   0
      Font.Color      =   -2147483635
      Font.Height     =   -29
      Font.Name       =   "Comic Sans MS"
      Font.Style      =   3
      ParentFont      =   0
      _Handle         =   1
   End
   Begin MMToolsX.MMLEDSpinX MMLEDSpinX1 
      Height          =   645
      Left            =   11520
      TabIndex        =   53
      Top             =   7080
      Width           =   1020
      Enabled         =   1
      LEDColor        =   65280
      InactiveColor   =   32768
      NumDigits       =   2
      DigitSize       =   12
      LEDSpace        =   1
      Increment       =   1
      ZeroBlank       =   0
      MaxValue        =   25
      MinValue        =   0
      Value           =   0
      DownGlyph.Data  =   $"Form1.frx":2E68
      DownNumGlyphs   =   3
      UpGlyph.Data    =   $"Form1.frx":3E32
      UpNumGlyphs     =   3
      ButtonFace      =   0
      MiddleButton    =   0
      _Handle         =   3
   End
   Begin MMToolsX.MMSliderX MMSliderX1 
      Height          =   1095
      Left            =   12120
      TabIndex        =   65
      Top             =   8280
      Width           =   255
      Object.Width           =   17
      Object.Height          =   73
      Color           =   -2147483633
      Enabled         =   0
      Bevel.BorderWidth=   5
      MaxValue        =   3
      Orientation     =   0
      ThumbWidth      =   23
      ThumbHeight     =   11
      Scale.Visible   =   0
      Scale.Color     =   0
      Scale.Style     =   0
      Scale.TickCount =   11
      Scale.EnlargeEvery=   5
      Scale.Size      =   7
      Scale.Origin    =   1
      Scale.Connect   =   1
      _Handle         =   6
   End
   Begin VB.CommandButton Command33 
      Caption         =   "enable speeds to be linked with wheel"
      Height          =   615
      Left            =   9120
      TabIndex        =   74
      Top             =   6840
      Width           =   1335
   End
   Begin VB.Timer Timer70 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   10440
      Top             =   8040
   End
   Begin VB.Timer Timer69 
      Interval        =   1
      Left            =   12480
      Top             =   8400
   End
   Begin VB.Timer Timer68 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   10440
      Top             =   8520
   End
   Begin VB.CommandButton Command32 
      Caption         =   "BUY NITROUS"
      Height          =   495
      Left            =   12600
      TabIndex        =   73
      Top             =   9480
      Width           =   1575
   End
   Begin VB.CommandButton Command31 
      Caption         =   "BUY NITROUS"
      Height          =   495
      Left            =   11040
      TabIndex        =   72
      Top             =   9480
      Width           =   1455
   End
   Begin VB.CommandButton Command30 
      Caption         =   "NITROUS!"
      Height          =   375
      Left            =   13200
      TabIndex        =   71
      Top             =   9120
      Width           =   975
   End
   Begin VB.CommandButton Command29 
      Cancel          =   -1  'True
      Caption         =   "NITROUS!"
      Height          =   375
      Left            =   11040
      TabIndex        =   70
      Top             =   9120
      Width           =   975
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   13200
      TabIndex        =   69
      Text            =   "100"
      Top             =   8760
      Width           =   975
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   11040
      TabIndex        =   68
      Text            =   "100"
      Top             =   8760
      Width           =   975
   End
   Begin VB.CommandButton Command28 
      Caption         =   "start both cars"
      BeginProperty Font 
         Name            =   "Marking Pen"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      TabIndex        =   63
      Top             =   9000
      Width           =   2175
   End
   Begin VB.Timer Timer67 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   12600
      Top             =   7080
   End
   Begin VB.CommandButton Command27 
      Caption         =   "submit"
      Height          =   375
      Left            =   14640
      TabIndex        =   62
      Top             =   8280
      Width           =   615
   End
   Begin VB.CommandButton Command13 
      Caption         =   "submit"
      Height          =   495
      Left            =   10920
      TabIndex        =   61
      Top             =   7800
      Width           =   1095
   End
   Begin VB.Timer Timer66 
      Interval        =   1
      Left            =   14280
      Top             =   1800
   End
   Begin VB.CommandButton Command26 
      Caption         =   "switch to day drive(this will restart the program"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   52
      Top             =   240
      Width           =   2415
   End
   Begin VB.CommandButton Command25 
      Caption         =   "change to night drive"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   51
      Top             =   240
      Width           =   2175
   End
   Begin VB.CommandButton Command24 
      BackColor       =   &H00E0E0E0&
      Caption         =   "help"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   24
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      MaskColor       =   &H000080FF&
      TabIndex        =   49
      Top             =   840
      Width           =   1335
   End
   Begin VB.Timer Timer65 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   10800
      Top             =   10560
   End
   Begin VB.Timer Timer64 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   10800
      Top             =   9960
   End
   Begin VB.Timer Timer63 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9720
      Top             =   10560
   End
   Begin VB.Timer Timer62 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9720
      Top             =   9960
   End
   Begin VB.CommandButton Command23 
      BackColor       =   &H00E0E0E0&
      Caption         =   "disable route2 status"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12720
      MaskColor       =   &H000080FF&
      TabIndex        =   44
      Top             =   10560
      Width           =   1335
   End
   Begin VB.CommandButton Command22 
      BackColor       =   &H00E0E0E0&
      Caption         =   "enable route2 status"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11280
      MaskColor       =   &H000080FF&
      TabIndex        =   43
      Top             =   10560
      Width           =   1335
   End
   Begin VB.CommandButton Command21 
      BackColor       =   &H00E0E0E0&
      Caption         =   "disable route1 status"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12720
      MaskColor       =   &H000080FF&
      TabIndex        =   42
      Top             =   10080
      Width           =   1335
   End
   Begin VB.CommandButton Command20 
      BackColor       =   &H00E0E0E0&
      Caption         =   "enable route1 status"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11280
      MaskColor       =   &H000080FF&
      TabIndex        =   41
      Top             =   10080
      Width           =   1335
   End
   Begin VB.Timer Timer61 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2760
      Top             =   10440
   End
   Begin VB.Timer Timer60 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2760
      Top             =   9840
   End
   Begin VB.Timer Timer59 
      Interval        =   1
      Left            =   11040
      Top             =   720
   End
   Begin VB.Timer Timer58 
      Interval        =   1
      Left            =   12840
      Top             =   1680
   End
   Begin VB.Timer Timer57 
      Interval        =   1
      Left            =   12840
      Top             =   1200
   End
   Begin VB.Timer Timer56 
      Interval        =   1
      Left            =   12360
      Top             =   1680
   End
   Begin VB.Timer Timer55 
      Interval        =   1
      Left            =   12360
      Top             =   1200
   End
   Begin VB.Timer Timer54 
      Interval        =   1
      Left            =   11880
      Top             =   1680
   End
   Begin VB.Timer Timer53 
      Interval        =   1
      Left            =   11400
      Top             =   1680
   End
   Begin VB.Timer Timer52 
      Interval        =   1
      Left            =   11880
      Top             =   1200
   End
   Begin VB.Timer Timer51 
      Interval        =   1
      Left            =   11400
      Top             =   1200
   End
   Begin VB.CommandButton Command19 
      BackColor       =   &H00E0E0E0&
      Caption         =   "end program"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   24
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   240
      MaskColor       =   &H000080FF&
      TabIndex        =   35
      Top             =   825
      Width           =   2295
   End
   Begin VB.CommandButton Command18 
      BackColor       =   &H00E0E0E0&
      Caption         =   "restart lights"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13320
      MaskColor       =   &H000080FF&
      TabIndex        =   34
      Top             =   6120
      Width           =   1935
   End
   Begin VB.CommandButton Command17 
      Caption         =   "remove debris"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11520
      TabIndex        =   33
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H00E0E0E0&
      Caption         =   "red light disable"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11040
      MaskColor       =   &H000080FF&
      TabIndex        =   32
      Top             =   6120
      Width           =   2175
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00E0E0E0&
      Caption         =   "remove all other cars except one call car to return to route2 start"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7800
      MaskColor       =   &H000080FF&
      TabIndex        =   30
      Top             =   7560
      Width           =   2535
   End
   Begin VB.Timer Timer50 
      Interval        =   1000
      Left            =   13800
      Top             =   1800
   End
   Begin VB.Timer Timer49 
      Interval        =   1
      Left            =   10080
      Top             =   4200
   End
   Begin VB.Timer Timer48 
      Interval        =   1
      Left            =   9600
      Top             =   4200
   End
   Begin VB.Timer Timer47 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   13680
      Top             =   4560
   End
   Begin VB.Timer Timer46 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   13680
      Top             =   3600
   End
   Begin VB.Timer Timer45 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   13680
      Top             =   2400
   End
   Begin VB.Timer Timer44 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   11640
      Top             =   4560
   End
   Begin VB.Timer Timer43 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   11640
      Top             =   3480
   End
   Begin VB.Timer Timer42 
      Interval        =   2000
      Left            =   11640
      Top             =   2400
   End
   Begin VB.Timer Timer41 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   7320
      Top             =   8400
   End
   Begin VB.Timer Timer40 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6840
      Top             =   8400
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H00E0E0E0&
      Caption         =   "no"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Marking Pen"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      MaskColor       =   &H000080FF&
      TabIndex        =   21
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H00E0E0E0&
      Caption         =   "yes"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Marking Pen"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      MaskColor       =   &H000080FF&
      TabIndex        =   20
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Timer Timer39 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6720
      Top             =   4680
   End
   Begin VB.Timer Timer38 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6240
      Top             =   4680
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00E0E0E0&
      Caption         =   "start route 2 cars"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Marking Pen"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      MaskColor       =   &H000080FF&
      TabIndex        =   19
      Top             =   8400
      Width           =   2535
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00E0E0E0&
      Caption         =   "start route 1 cars"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Marking Pen"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      MaskColor       =   &H000080FF&
      TabIndex        =   18
      Top             =   8400
      Width           =   2415
   End
   Begin VB.Timer Timer37 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   7320
      Top             =   7800
   End
   Begin VB.Timer Timer36 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6840
      Top             =   7800
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00E0E0E0&
      Caption         =   "remove all other cars except one to return to route1 start"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4320
      MaskColor       =   &H000080FF&
      TabIndex        =   17
      Top             =   7560
      Width           =   2415
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   1455
      Left            =   6480
      TabIndex        =   2
      Top             =   120
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   2566
      _Version        =   393216
      Tabs            =   6
      Tab             =   4
      TabHeight       =   520
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "route1.1"
      TabPicture(0)   =   "Form1.frx":4DFC
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Command3"
      Tab(0).Control(1)=   "Text1"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "route1.2"
      TabPicture(1)   =   "Form1.frx":4E18
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command4"
      Tab(1).Control(1)=   "Text2"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "route1.3"
      TabPicture(2)   =   "Form1.frx":4E34
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Text3"
      Tab(2).Control(1)=   "Command5"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "route2.1"
      TabPicture(3)   =   "Form1.frx":4E50
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Command6"
      Tab(3).Control(1)=   "Text4"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "route2.2"
      TabPicture(4)   =   "Form1.frx":4E6C
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Text5"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Command7"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "route2.3"
      TabPicture(5)   =   "Form1.frx":4E88
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Text6"
      Tab(5).Control(1)=   "Command8"
      Tab(5).ControlCount=   2
      Begin VB.CommandButton Command8 
         Caption         =   "ok"
         Height          =   495
         Left            =   -72360
         TabIndex        =   14
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton Command7 
         Caption         =   "ok"
         Height          =   495
         Left            =   2280
         TabIndex        =   13
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton Command6 
         Caption         =   "ok"
         Height          =   495
         Left            =   -72720
         TabIndex        =   12
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "ok"
         Height          =   495
         Left            =   -72360
         TabIndex        =   11
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "ok"
         Height          =   495
         Left            =   -72240
         TabIndex        =   10
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "ok"
         Height          =   495
         Left            =   -72600
         TabIndex        =   9
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox Text6 
         Height          =   495
         Left            =   -74520
         TabIndex        =   8
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox Text5 
         Height          =   615
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox Text4 
         Height          =   615
         Left            =   -74640
         TabIndex        =   6
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Height          =   495
         Left            =   -74640
         TabIndex        =   5
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   -74640
         TabIndex        =   4
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   -74640
         TabIndex        =   3
         Top             =   720
         Width           =   1695
      End
   End
   Begin VB.Timer Timer35 
      Interval        =   1
      Left            =   5160
      Top             =   4680
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "create route 2 cars"
      BeginProperty Font 
         Name            =   "Marking Pen"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      MaskColor       =   &H000080FF&
      TabIndex        =   1
      Top             =   7680
      Width           =   1695
   End
   Begin VB.Timer Timer34 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   6840
      Top             =   2040
   End
   Begin VB.Timer Timer33 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   6360
      Top             =   2040
   End
   Begin VB.Timer Timer32 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   5880
      Top             =   2040
   End
   Begin VB.Timer Timer31 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   5400
      Top             =   2040
   End
   Begin VB.Timer Timer30 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   4920
      Top             =   2040
   End
   Begin VB.Timer Timer29 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   4440
      Top             =   2040
   End
   Begin VB.Timer Timer28 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   3960
      Top             =   2040
   End
   Begin VB.Timer Timer27 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   3480
      Top             =   2040
   End
   Begin VB.Timer Timer26 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   3000
      Top             =   2040
   End
   Begin VB.Timer Timer25 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   2520
      Top             =   2040
   End
   Begin VB.Timer Timer24 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   2040
      Top             =   2040
   End
   Begin VB.Timer Timer23 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   1560
      Top             =   2040
   End
   Begin VB.Timer Timer22 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   1080
      Top             =   2040
   End
   Begin VB.Timer Timer21 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   600
      Top             =   2040
   End
   Begin VB.Timer Timer20 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   120
      Top             =   2040
   End
   Begin VB.Timer Timer19 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   4920
      Top             =   1560
   End
   Begin VB.Timer Timer18 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   4440
      Top             =   1560
   End
   Begin VB.Timer Timer17 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   3960
      Top             =   1560
   End
   Begin VB.Timer Timer16 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   3480
      Top             =   1560
   End
   Begin VB.Timer Timer15 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   3000
      Top             =   1560
   End
   Begin VB.Timer Timer14 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   2520
      Top             =   1560
   End
   Begin VB.Timer Timer13 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   2040
      Top             =   1560
   End
   Begin VB.Timer Timer12 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   1560
      Top             =   1560
   End
   Begin VB.Timer Timer11 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   1080
      Top             =   1560
   End
   Begin VB.Timer Timer10 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   600
      Top             =   1560
   End
   Begin VB.Timer Timer9 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   120
      Top             =   1560
   End
   Begin VB.Timer Timer8 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   3960
      Top             =   8520
   End
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   3480
      Top             =   8520
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   3000
      Top             =   8520
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   2520
      Top             =   8520
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "create route 1 cars"
      BeginProperty Font 
         Name            =   "Marking Pen"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      MaskColor       =   &H000080FF&
      TabIndex        =   0
      Top             =   7680
      Width           =   1575
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   1800
      Top             =   8520
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   1320
      Top             =   8520
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   840
      Top             =   8520
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   360
      Top             =   8520
   End
   Begin VB.Label Label26 
      Caption         =   "car2 nitrous"
      Height          =   255
      Left            =   13200
      TabIndex        =   66
      Top             =   8400
      Width           =   975
   End
   Begin VB.Label Label25 
      Caption         =   "car1 nitrous"
      Height          =   255
      Left            =   11040
      TabIndex        =   64
      Top             =   8400
      Width           =   975
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "car2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   14400
      TabIndex        =   60
      Top             =   7080
      Width           =   735
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "car1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   13320
      TabIndex        =   59
      Top             =   7080
      Width           =   855
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "max speeds"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   13800
      TabIndex        =   58
      Top             =   6720
      Width           =   735
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "laps"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   10560
      TabIndex        =   57
      Top             =   7080
      Width           =   975
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "make them race!"
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
      Height          =   615
      Left            =   10920
      TabIndex        =   54
      Top             =   6600
      Width           =   3015
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "route2 status"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   360
      TabIndex        =   50
      Top             =   10440
      Width           =   2295
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   10440
      TabIndex        =   48
      Top             =   10560
      Width           =   1095
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   10440
      TabIndex        =   47
      Top             =   9960
      Width           =   975
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Laps"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   9240
      TabIndex        =   46
      Top             =   10560
      Width           =   1095
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Laps"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   9240
      TabIndex        =   45
      Top             =   9960
      Width           =   1095
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "route1 status"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   360
      TabIndex        =   39
      Top             =   9840
      Width           =   2895
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Intersection"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   37
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Underpass bridge"
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
      Left            =   1680
      TabIndex        =   36
      Top             =   6480
      Width           =   4455
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Disable the red lights to watch them BANG!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10800
      TabIndex        =   31
      Top             =   5280
      Width           =   3495
   End
   Begin VB.Label Label10 
      Caption         =   "yellow"
      Height          =   975
      Left            =   14160
      TabIndex        =   29
      Top             =   4440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "red"
      Height          =   735
      Left            =   14160
      TabIndex        =   28
      Top             =   3360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "green"
      Height          =   735
      Left            =   14280
      TabIndex        =   27
      Top             =   2160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "yellow"
      Height          =   975
      Left            =   11880
      TabIndex        =   26
      Top             =   4440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "red"
      Height          =   855
      Left            =   12000
      TabIndex        =   25
      Top             =   3240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "green"
      Height          =   735
      Left            =   12000
      TabIndex        =   24
      Top             =   2160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Shape42 
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   12480
      Shape           =   3  'Circle
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Shape Shape41 
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   12480
      Shape           =   3  'Circle
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Shape Shape40 
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   12480
      Shape           =   3  'Circle
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Shape Shape39 
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   10440
      Shape           =   3  'Circle
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Shape Shape38 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   12480
      Shape           =   3  'Circle
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Shape Shape37 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   12480
      Shape           =   3  'Circle
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Shape Shape36 
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   12480
      Shape           =   3  'Circle
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Shape Shape35 
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   10440
      Shape           =   3  'Circle
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Shape Shape34 
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   10440
      Shape           =   3  'Circle
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Shape Shape33 
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   10440
      Shape           =   3  'Circle
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Shape Shape32 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   10440
      Shape           =   3  'Circle
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Shape Shape31 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   10440
      Shape           =   3  'Circle
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "set to default:NO"
      BeginProperty Font 
         Name            =   "HL2cross"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   6360
      TabIndex        =   23
      Top             =   3600
      Width           =   4335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "use safety control"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   22
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Line Line24 
      X1              =   8640
      X2              =   6120
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line23 
      X1              =   6120
      X2              =   7800
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "route1"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   480
      TabIndex        =   16
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "route2"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   840
      TabIndex        =   15
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Line Line19 
      X1              =   1680
      X2              =   720
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Shape Shape28 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   5
      FillColor       =   &H000000FF&
      FillStyle       =   7  'Diagonal Cross
      Height          =   1455
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   3480
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Line Line18 
      X1              =   2280
      X2              =   2280
      Y1              =   4680
      Y2              =   6480
   End
   Begin VB.Line Line17 
      X1              =   1680
      X2              =   1680
      Y1              =   4680
      Y2              =   6480
   End
   Begin VB.Line Line16 
      X1              =   6120
      X2              =   6120
      Y1              =   4680
      Y2              =   6480
   End
   Begin VB.Line Line15 
      X1              =   5520
      X2              =   5520
      Y1              =   4680
      Y2              =   6480
   End
   Begin VB.Line Line14 
      X1              =   5520
      X2              =   5520
      Y1              =   3360
      Y2              =   4080
   End
   Begin VB.Line Line13 
      X1              =   2280
      X2              =   5520
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line12 
      X1              =   2280
      X2              =   2280
      Y1              =   4080
      Y2              =   3360
   End
   Begin VB.Line Line11 
      X1              =   6120
      X2              =   6120
      Y1              =   2520
      Y2              =   4080
   End
   Begin VB.Line Line10 
      X1              =   720
      X2              =   6120
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line9 
      X1              =   1680
      X2              =   1680
      Y1              =   4080
      Y2              =   3240
   End
   Begin VB.Line Line8 
      X1              =   480
      X2              =   480
      Y1              =   4680
      Y2              =   7320
   End
   Begin VB.Line Line7 
      X1              =   8640
      X2              =   480
      Y1              =   7320
      Y2              =   7320
   End
   Begin VB.Line Line6 
      X1              =   1320
      X2              =   1320
      Y1              =   4680
      Y2              =   6480
   End
   Begin VB.Line Line5 
      X1              =   7800
      X2              =   1320
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line Line4 
      X1              =   7800
      X2              =   7800
      Y1              =   4680
      Y2              =   6480
   End
   Begin VB.Line Line3 
      X1              =   8640
      X2              =   8640
      Y1              =   4080
      Y2              =   7320
   End
   Begin VB.Line Line2 
      X1              =   480
      X2              =   5520
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line1 
      X1              =   1320
      X2              =   5520
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Shape Shape30 
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   720
      Shape           =   4  'Rounded Rectangle
      Top             =   4920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape Shape29 
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   8040
      Shape           =   4  'Rounded Rectangle
      Top             =   5640
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape Shape27 
      FillColor       =   &H0080FF80&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   4800
      Shape           =   4  'Rounded Rectangle
      Top             =   2760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Shape Shape26 
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   7320
      Shape           =   4  'Rounded Rectangle
      Top             =   6720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Shape Shape25 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   6240
      Shape           =   4  'Rounded Rectangle
      Top             =   6720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Shape Shape24 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00404000&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   5040
      Shape           =   4  'Rounded Rectangle
      Top             =   6720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Shape Shape23 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00404000&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3960
      Shape           =   4  'Rounded Rectangle
      Top             =   6720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Shape Shape22 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00404000&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3000
      Shape           =   4  'Rounded Rectangle
      Top             =   6720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Shape Shape21 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00404000&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   1920
      Shape           =   4  'Rounded Rectangle
      Top             =   6720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Shape Shape20 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   720
      Shape           =   4  'Rounded Rectangle
      Top             =   6120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape Shape19 
      BackStyle       =   1  'Opaque
      FillColor       =   &H0080FF80&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3480
      Shape           =   4  'Rounded Rectangle
      Top             =   2760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Shape Shape18 
      BackStyle       =   1  'Opaque
      FillColor       =   &H0080FF80&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   2280
      Shape           =   4  'Rounded Rectangle
      Top             =   2760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Shape Shape17 
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   8040
      Shape           =   4  'Rounded Rectangle
      Top             =   4560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape Shape12 
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   7080
      Shape           =   4  'Rounded Rectangle
      Top             =   4200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Shape Shape11 
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   6120
      Shape           =   4  'Rounded Rectangle
      Top             =   4200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Shape Shape16 
      FillColor       =   &H0080FF80&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   5640
      Shape           =   4  'Rounded Rectangle
      Top             =   5760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape Shape15 
      FillColor       =   &H0080FF80&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   5640
      Shape           =   4  'Rounded Rectangle
      Top             =   3360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape Shape14 
      FillColor       =   &H0080FF80&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   5640
      Shape           =   4  'Rounded Rectangle
      Top             =   4200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape Shape10 
      FillColor       =   &H0080FF80&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   5640
      Shape           =   4  'Rounded Rectangle
      Top             =   5040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape Shape13 
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   5040
      Shape           =   4  'Rounded Rectangle
      Top             =   4200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Shape Shape9 
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   4200
      Shape           =   4  'Rounded Rectangle
      Top             =   4200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Shape Shape8 
      FillColor       =   &H0080FF80&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   1800
      Shape           =   4  'Rounded Rectangle
      Top             =   4680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape Shape7 
      FillColor       =   &H0080FF80&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   1800
      Shape           =   4  'Rounded Rectangle
      Top             =   3240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H0080FF80&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   1800
      Shape           =   4  'Rounded Rectangle
      Top             =   5760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H0080FF80&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   1800
      Shape           =   4  'Rounded Rectangle
      Top             =   5280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape Shape6 
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   4200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3240
      Shape           =   4  'Rounded Rectangle
      Top             =   4200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Shape Shape4 
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   2280
      Shape           =   4  'Rounded Rectangle
      Top             =   4200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   4200
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Timer1.Enabled = True
Command9.Enabled = True
End Sub

Private Sub Command10_Click()
Timer37.Enabled = True
Command2.Enabled = False
Timer41.Enabled = True
If Shape7.Visible = True Then
Command12.Enabled = True
End If
End Sub

Private Sub Command11_Click()
Timer36.Enabled = False
Timer1.Enabled = True
Command11.Enabled = False
Command1.Enabled = True
Timer40.Enabled = False
End Sub

Private Sub Command12_Click()
Timer37.Enabled = False
Timer20.Enabled = True
Command12.Enabled = False
Command2.Enabled = True
Timer41.Enabled = False
End Sub

Private Sub Command13_Click()
If MMLEDSpinX1.Value = "0" Then
MsgBox "enter value!"
Else
Timer67.Enabled = True
End If
End Sub

Private Sub Command14_Click()
Timer38.Enabled = True
Label4.Caption = "safety control is set to:YES"
End Sub

Private Sub Command15_Click()
Timer38.Enabled = False
Label4.Caption = "safety control is set to : NO"
End Sub

Private Sub Command16_Click()
Timer46.Interval = 1
Timer43.Interval = 1
Timer44.Interval = 1
Timer47.Interval = 1
Command14.Enabled = True
Command15.Enabled = True
End Sub

Private Sub Command17_Click()
Shape13.Visible = False
Shape28.Visible = False
Shape14.Visible = False
Shape15.Visible = False
Shape9.Visible = False
End Sub

Private Sub Command18_Click()
Timer46.Interval = 4000
Timer43.Interval = 4000
Timer44.Interval = 1000
Timer47.Interval = 1000
Timer44.Enabled = False
Timer43.Enabled = False
Timer42.Enabled = True
Timer45.Enabled = False
Timer46.Enabled = False
Timer47.Enabled = False
Timer50.Enabled = True
Command14.Enabled = False
Command15.Enabled = False
Timer38.Enabled = False
Label4.Caption = "set to default:NO"
End Sub

Private Sub Command19_Click()
End
End Sub

Private Sub Command2_Click()
Timer20.Enabled = True
Command10.Enabled = True
End Sub

Private Sub Command20_Click()
Timer60.Enabled = True
Timer62.Enabled = True
Timer64.Enabled = True
End Sub

Private Sub Command21_Click()
Timer60.Enabled = False
MMGaugeX1.Progress = 0
Timer62.Enabled = False
Label18.Caption = "0"
Timer64.Enabled = False
End Sub

Private Sub Command22_Click()
Timer61.Enabled = True
Timer63.Enabled = True
Timer65.Enabled = True
End Sub

Private Sub Command23_Click()
Timer61.Enabled = False
MMGaugeX2.Progress = 0
Timer63.Enabled = False
Timer65.Enabled = False
Label19.Caption = "0"
End Sub

Private Sub Command24_Click()
Form1.Hide
Form3.Show
End Sub

Private Sub Command25_Click()
Form1.BackColor = blue
End Sub

Private Sub Command26_Click()
Unload Form1
Load Form1
Form1.Show
End Sub

Private Sub Command27_Click()
Timer1.Interval = MMWheelX1.Value * 1
Timer2.Interval = MMWheelX1.Value * 700
Timer3.Interval = MMWheelX1.Value * 600
Timer4.Interval = MMWheelX1.Value * 500
Timer5.Interval = MMWheelX1.Value * 400
Timer6.Interval = MMWheelX1.Value * 400
Timer7.Interval = MMWheelX1.Value * 400
Timer8.Interval = MMWheelX1.Value * 400
Timer9.Interval = MMWheelX1.Value * 900
Timer10.Interval = MMWheelX1.Value * 900
Timer11.Interval = MMWheelX1.Value * 900
Timer12.Interval = MMWheelX1.Value * 800
Timer13.Interval = MMWheelX1.Value * 500
Timer14.Interval = MMWheelX1.Value * 200
Timer15.Interval = MMWheelX1.Value * 200
Timer16.Interval = MMWheelX1.Value * 200
Timer17.Interval = MMWheelX1.Value * 400
Timer18.Interval = MMWheelX1.Value * 800
Timer19.Interval = MMWheelX1.Value * 500
Timer20.Interval = MMWheelX2.Value * 1000
Timer21.Interval = MMWheelX2.Value * 900
Timer22.Interval = MMWheelX2.Value * 800
Timer23.Interval = MMWheelX2.Value * 700
Timer24.Interval = MMWheelX2.Value * 600
Timer25.Interval = MMWheelX2.Value * 600
Timer26.Interval = MMWheelX2.Value * 600
Timer27.Interval = MMWheelX2.Value * 800
Timer28.Interval = MMWheelX2.Value * 700
Timer29.Interval = MMWheelX2.Value * 600
Timer30.Interval = MMWheelX2.Value * 500
Timer31.Interval = MMWheelX2.Value * 400
Timer32.Interval = MMWheelX2.Value * 200
Timer33.Interval = MMWheelX2.Value * 500
Timer34.Interval = MMWheelX2.Value * 3000
End Sub

Private Sub Command28_Click()
Timer1.Enabled = True
Command9.Enabled = True
Timer20.Enabled = True
Command10.Enabled = True
End Sub

Private Sub Command29_Click()
Command29.Enabled = True
If MMSliderX1.Position > 0 Then
MMSliderX1.Position = Val(MMSliderX1.Position) - 1
Timer1.Interval = 100
Timer2.Interval = 100
Timer3.Interval = 100
Timer4.Interval = 100
Timer5.Interval = 100
Timer6.Interval = 100
Timer7.Interval = 100
Timer8.Interval = 100
Else
MsgBox "OUT OF NITROUS"
End If
End Sub

Private Sub Command3_Click()
Timer1.Interval = Text1.Text * 1000
Timer2.Interval = Text1.Text * 1000
Timer3.Interval = Text1.Text * 1000
Timer4.Interval = Text1.Text * 1000
Timer5.Interval = Text1.Text * 1000
Timer6.Interval = Text1.Text * 1000
End Sub

Private Sub Command30_Click()
Command30.Enabled = True
If MMSliderX2.Position > 0 Then
MMSliderX2.Position = Val(MMSliderX2.Position) - 1
Timer20.Interval = 100
Timer21.Interval = 100
Timer22.Interval = 100
Timer23.Interval = 100
Timer24.Interval = 100
Timer25.Interval = 100
Else
MsgBox "OUT OF NITROUS"
End If
End Sub

Private Sub Command31_Click()
Text7.Text = Val(Text7.Text) - 70
If Val(Text7.Text) > 0 Then
MMSliderX1.Position = Val(MMSliderX1.Position) + 1
Else
MsgBox "not enough funds"
Text7.Text = Val(Text7.Text) + 70
End If
End Sub

Private Sub Command32_Click()
Text8.Text = Val(Text8.Text) - 70
If Val(Text8.Text) > 0 Then
MMSliderX2.Position = Val(MMSliderX2.Position) + 1
Else
MsgBox "not enough funds"
Text8.Text = Val(Text8.Text) + 70
End If
End Sub

Private Sub Command33_Click()
Command33.Enabled = False
Timer68.Enabled = True
Timer70.Enabled = True
End Sub

Private Sub Command4_Click()
Timer7.Interval = Text2.Text * 1000
Timer8.Interval = Text2.Text * 1000
Timer9.Interval = Text2.Text * 1000
Timer10.Interval = Text2.Text * 1000
Timer11.Interval = Text2.Text * 1000
Timer12.Interval = Text2.Text * 1000
End Sub

Private Sub Command5_Click()
Timer13.Interval = Text3.Text * 1000
Timer14.Interval = Text3.Text * 1000
Timer15.Interval = Text3.Text * 1000
Timer16.Interval = Text3.Text * 1000
Timer17.Interval = Text3.Text * 1000
Timer18.Interval = Text3.Text * 1000
Timer19.Interval = Text3.Text * 1000
End Sub

Private Sub Command6_Click()
Timer20.Interval = Text4.Text * 1000
Timer21.Interval = Text4.Text * 1000
Timer22.Interval = Text4.Text * 1000
Timer23.Interval = Text4.Text * 1000
Timer24.Interval = Text4.Text * 1000
Timer25.Interval = Text4.Text * 1000
End Sub

Private Sub Command7_Click()
Timer26.Interval = Text5.Text * 1000
Timer27.Interval = Text5.Text * 1000
Timer28.Interval = Text5.Text * 1000
Timer29.Interval = Text5.Text * 1000
Timer30.Interval = Text5.Text * 1000
End Sub

Private Sub Command8_Click()
Timer31.Interval = Text6.Text * 1000
Timer32.Interval = Text6.Text * 1000
Timer33.Interval = Text6.Text * 1000
Timer34.Interval = Text6.Text * 1000
End Sub

Private Sub Command9_Click()
Timer36.Enabled = True
Command1.Enabled = False
Timer40.Enabled = True
If Shape30.Visible = True Then
Command11.Enabled = True
End If
End Sub

Private Sub SSTab2_DblClick()

End Sub

Private Sub Form_Load()
MsgBox "verify yourself"
End Sub

Private Sub SoundRec1_GotFocus()

End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
Shape1.Visible = True
Timer2.Enabled = True
Shape30.Visible = False
Shape8.Visible = False
End Sub

Private Sub Timer17_Timer()
Timer17.Enabled = False
Shape20.Visible = True
Shape21.Visible = False
Timer18.Enabled = True
End Sub

Private Sub Timer10_Timer()
Timer10.Enabled = False
Shape17.Visible = False
Shape29.Visible = True
Timer11.Enabled = True
End Sub

Private Sub Timer11_Timer()
Timer11.Enabled = False
Shape26.Visible = True
Shape29.Visible = False
Timer12.Enabled = True
End Sub

Private Sub Timer12_Timer()
Timer12.Enabled = False
Shape25.Visible = True
Shape26.Visible = False
Timer13.Enabled = True
End Sub

Private Sub Timer13_Timer()
Timer13.Enabled = False
Shape24.Visible = True
Shape25.Visible = False
Timer14.Enabled = True
End Sub

Private Sub Timer14_Timer()
Timer14.Enabled = False
Shape23.Visible = True
Shape24.Visible = False
Timer15.Enabled = True
End Sub

Private Sub Timer15_Timer()
Timer15.Enabled = False
Shape22.Visible = True
Shape23.Visible = False
Timer16.Enabled = True
End Sub

Private Sub Timer16_Timer()
Timer16.Enabled = False
Shape21.Visible = True
Shape22.Visible = False
Timer17.Enabled = True
End Sub

Private Sub Timer18_Timer()
Timer18.Enabled = False
Shape30.Visible = True
Shape20.Visible = False
Timer19.Enabled = True
End Sub

Private Sub Timer19_Timer()
Timer19.Enabled = False
Shape30.Visible = True
Shape20.Visible = False
Timer1.Enabled = True
End Sub

Private Sub Timer2_Timer()
Timer2.Enabled = False
Shape1.Visible = False
Shape6.Visible = True
Timer3.Enabled = True
End Sub

Private Sub Timer20_Timer()
Timer20.Enabled = False
Shape18.Visible = True
Timer21.Enabled = True
Shape8.Visible = False
Shape7.Visible = False
End Sub

Private Sub Timer21_Timer()
Timer21.Enabled = False
Shape19.Visible = True
Shape18.Visible = False
Timer22.Enabled = True
End Sub

Private Sub Timer22_Timer()
Timer22.Enabled = False
Shape27.Visible = True
Shape19.Visible = False
Timer23.Enabled = True
End Sub

Private Sub Timer23_Timer()
Timer23.Enabled = False
Shape15.Visible = True
Shape27.Visible = False
Timer24.Enabled = True
End Sub

Private Sub Timer24_Timer()
Timer24.Enabled = False
Shape14.Visible = True
Shape15.Visible = False
Timer25.Enabled = True
End Sub

Private Sub Timer25_Timer()
Timer25.Enabled = False
Shape10.Visible = True
Shape14.Visible = False
Timer26.Enabled = True
End Sub

Private Sub Timer26_Timer()
Timer26.Enabled = False
Shape16.Visible = True
Shape10.Visible = False
Timer27.Enabled = True
End Sub

Private Sub Timer27_Timer()
Timer27.Enabled = False
Shape24.Visible = True
Shape16.Visible = False
Timer28.Enabled = True
End Sub

Private Sub Timer28_Timer()
Timer28.Enabled = False
Shape23.Visible = True
Shape24.Visible = False
Timer29.Enabled = True
End Sub

Private Sub Timer29_Timer()
Timer29.Enabled = False
Shape22.Visible = True
Shape23.Visible = False
Timer30.Enabled = True
End Sub

Private Sub Timer3_Timer()
Timer3.Enabled = False
Timer4.Enabled = True
Shape6.Visible = False
Shape4.Visible = True
End Sub

Private Sub Timer30_Timer()
Timer30.Enabled = False
Shape21.Visible = True
Shape22.Visible = False
Timer31.Enabled = True
End Sub

Private Sub Timer31_Timer()
Timer31.Enabled = False
Shape3.Visible = True
Shape21.Visible = False
Timer32.Enabled = True
End Sub

Private Sub Timer32_Timer()
Timer32.Enabled = False
Shape2.Visible = True
Shape3.Visible = False
Timer33.Enabled = True
End Sub

Private Sub Timer33_Timer()
Timer33.Enabled = False
Shape8.Visible = True
Shape2.Visible = False
Timer34.Enabled = True
End Sub

Private Sub Timer34_Timer()
Timer34.Enabled = False
Shape7.Visible = True
Shape2.Visible = False
Shape8.Visible = False
Timer20.Enabled = True
End Sub

Private Sub Timer35_Timer()
If Shape13.Visible = True And Shape13.Left = 5040 And Shape14.Visible = True And Shape14.Left = 5640 Then
Shape28.Visible = True
Timer7.Enabled = False
Timer25.Enabled = False
End If
End Sub

Private Sub Timer36_Timer()
If Timer1.Enabled = True Then
Timer1.Enabled = False
End If
End Sub

Private Sub Timer37_Timer()
If Timer20.Enabled = True Then
Timer20.Enabled = False
End If
End Sub

Private Sub Timer38_Timer()
If Shape15.Visible = True And Shape9.Visible = True Then
Timer6.Enabled = False
Timer39.Enabled = True
Else
If Shape15.Visible = True And Shape14.Visible = True Then
Timer6.Enabled = False
Timer39.Enabled = True
End If
End If
End Sub

Private Sub Timer39_Timer()
If Shape9.Visible = True And Shape10.Visible = True Then
Timer39.Enabled = False
Timer6.Enabled = True
Else
Timer6.Enabled = False
End If
End Sub

Private Sub Timer4_Timer()
Timer4.Enabled = False
Timer5.Enabled = True
Shape4.Visible = False
Shape5.Visible = True
End Sub

Private Sub Timer40_Timer()
If Shape30.Visible = True Then
Command11.Enabled = True
End If
End Sub

Private Sub Timer41_Timer()
If Shape7.Visible = True Then
Command12.Enabled = True
End If
End Sub

Private Sub Timer42_Timer()
Timer42.Enabled = False
Shape33.Visible = False
Shape39.Visible = True
Timer43.Enabled = True
End Sub

Private Sub Timer43_Timer()
Timer43.Enabled = False
Timer44.Enabled = True
Shape34.Visible = False
End Sub

Private Sub Timer44_Timer()
Timer44.Enabled = False
Timer42.Enabled = True
Shape34.Visible = True
Shape39.Visible = False
Shape33.Visible = True
End Sub

Private Sub Timer45_Timer()
Timer45.Enabled = False
Timer46.Enabled = True
Shape40.Visible = False
Shape42.Visible = True
End Sub

Private Sub Timer46_Timer()
Timer46.Enabled = False
Timer47.Enabled = True
Shape41.Visible = False
End Sub

Private Sub Timer47_Timer()
Timer47.Enabled = False
Timer45.Enabled = True
Shape40.Visible = True
Shape41.Visible = True
Shape42.Visible = False
End Sub

Private Sub Timer48_Timer()
If Shape33.Visible = False Then
Timer6.Enabled = False
Else
If Shape9.Visible = True Then
Timer6.Enabled = True
End If
End If
End Sub

Private Sub Timer49_Timer()
If Shape40.Visible = False Then
Timer24.Enabled = False
Else
If Shape15.Visible = True Then
Timer24.Enabled = True
End If
End If
End Sub

Private Sub Timer5_Timer()
Timer5.Enabled = False
Shape9.Visible = True
Timer6.Enabled = True
Shape12.Visible = False
Shape5.Visible = False
End Sub

Private Sub Timer50_Timer()
Timer47.Enabled = True
Timer50.Enabled = False
End Sub

Private Sub Timer51_Timer()
If Shape9.Visible = True And Shape5.Visible = True Then
Timer5.Enabled = False
Else
If Shape5.Visible = True Then
Timer5.Enabled = True
End If
End If
End Sub

Private Sub Timer52_Timer()
If Shape28.Visible = True And Shape13.Visible = True Then
Timer6.Enabled = False
End If
End Sub

Private Sub Timer53_Timer()
If Shape28.Visible = True And Shape15.Visible = True Then
Timer24.Enabled = False
End If
End Sub

Private Sub Timer54_Timer()
If Shape27.Visible = True And Shape15.Visible = True Then
Timer23.Enabled = False
Else
If Shape27.Visible = True Then
Timer23.Enabled = True
End If
End If
End Sub

Private Sub Timer55_Timer()
If Shape19.Visible = True And Shape27.Visible = True Then
Timer22.Enabled = False
Else
If Shape19.Visible = True Then
Timer22.Enabled = True
End If
End If
End Sub

Private Sub Timer56_Timer()
If Shape4.Visible = True And Shape5.Visible = True Then
Timer4.Enabled = False
Else
If Shape4.Visible = True Then
Timer4.Enabled = True
End If
End If
End Sub

Private Sub Timer57_Timer()
If Shape19.Visible = True And Shape18.Visible = True Then
Timer21.Enabled = False
Else
If Shape18.Visible = True Then
Timer21.Enabled = True
End If
End If
End Sub

Private Sub Timer58_Timer()
If Shape6.Visible = True And Shape4.Visible = True Then
Timer3.Enabled = False
Else
If Shape6.Visible = True Then
Timer3.Enabled = True
End If
End If
End Sub

Private Sub Timer59_Timer()
If Shape28.Visible = True Then
Command17.Enabled = True
Else
Command17.Enabled = False
End If
End Sub

Private Sub Timer6_Timer()
Timer6.Enabled = False
Shape9.Visible = False
Shape13.Visible = True
Timer7.Enabled = True
End Sub

Private Sub Timer60_Timer()
If Shape1.Visible = True Then
MMGaugeX1.Progress = 6
Else
If Shape6.Visible = True Then
MMGaugeX1.Progress = 12
Else
If Shape4.Visible = True Then
MMGaugeX1.Progress = 17
Else
If Shape5.Visible = True Then
MMGaugeX1.Progress = 23
Else
If Shape9.Visible = True Then
MMGaugeX1.Progress = 27
Else
If Shape13.Visible = True Then
MMGaugeX1.Progress = 32
Else
If Shape11.Visible = True Then
MMGaugeX1.Progress = 38
Else
If Shape12.Visible = True Then
MMGaugeX1.Progress = 44
Else
If Shape17.Visible = True Then
MMGaugeX1.Progress = 50
Else
If Shape29.Visible = True Then
MMGaugeX1.Progress = 56
Else
If Shape26.Visible = True Then
MMGaugeX1.Progress = 62
Else
If Shape25.Visible = True Then
MMGaugeX1.Progress = 68
Else
If Shape24.Visible = True And Timer14.Enabled = True Then
MMGaugeX1.Progress = 74
Else
If Shape23.Visible = True And Timer15.Enabled = True Then
MMGaugeX1.Progress = 80
Else
If Shape22.Visible = True And Timer16.Enabled = True Then
MMGaugeX1.Progress = 85
Else
If Shape21.Visible = True And Timer17.Enabled = True Then
MMGaugeX1.Progress = 90
Else
If Shape20.Visible = True Then
MMGaugeX1.Progress = 95
Else
If Shape30.Visible = True Then
MMGaugeX1.Progress = 100
Else
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End Sub

Private Sub Timer61_Timer()
If Shape18.Visible = True Then
MMGaugeX2.Progress = 7
Else
If Shape19.Visible = True Then
MMGaugeX2.Progress = 14
Else
If Shape27.Visible = True Then
MMGaugeX2.Progress = 21
Else
If Shape15.Visible = True Then
MMGaugeX2.Progress = 28
Else
If Shape14.Visible = True Then
MMGaugeX2.Progress = 35
Else
If Shape10.Visible = True Then
MMGaugeX2.Progress = 42
Else
If Shape16.Visible = True Then
MMGaugeX2.Progress = 49
Else
If Shape24.Visible = True And Timer28.Enabled = True Then
MMGaugeX2.Progress = 56
Else
If Shape23.Visible = True And Timer29.Enabled = True Then
MMGaugeX2.Progress = 63
Else
If Shape22.Visible = True And Timer30.Enabled = True Then
MMGaugeX2.Progress = 70
Else
If Shape21.Visible = True And Timer31.Enabled = True Then
MMGaugeX2.Progress = 77
Else
If Shape3.Visible = True Then
MMGaugeX2.Progress = 83
Else
If Shape2.Visible = True Then
MMGaugeX2.Progress = 89
Else
If Shape8.Visible = True Then
MMGaugeX2.Progress = 95
Else
If Shape7.Visible = True Then
MMGaugeX2.Progress = 100
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End Sub

Private Sub Timer62_Timer()
If Shape30.Visible = True Then
Label18.Caption = Val(Label18.Caption) + 1
Timer62.Enabled = False
End If
End Sub

Private Sub Timer63_Timer()
If MMGaugeX2.Progress = 100 Then
Label19.Caption = Val(Label19.Caption) + 1
Timer63.Enabled = False
End If
End Sub

Private Sub Timer64_Timer()
If Shape1.Visible = True Then
Timer62.Enabled = True
End If
End Sub

Private Sub Timer65_Timer()
If Shape18.Visible = True Then
Timer63.Enabled = True
End If
End Sub

Private Sub Timer66_Timer()
If Shape39.Visible = False And Shape42.Visible = False Then
Timer44.Enabled = False
Timer43.Enabled = False
Timer42.Enabled = True
Timer45.Enabled = False
Timer46.Enabled = False
Timer47.Enabled = False
Timer50.Enabled = True
End If
End Sub

Private Sub Timer67_Timer()
If Label18.Caption = MMLEDSpinX1.Value Then
MsgBox "car1 wins!"
Timer67.Enabled = False
Timer19.Enabled = False
Text7.Text = Val(Text7.Text) + 100
Text8.Text = Val(Text8.Text) + 20
Timer37.Enabled = True
Timer41.Enabled = True
Else
If Label19.Caption = MMLEDSpinX1.Value Then
MsgBox "car2 wins!"
Timer67.Enabled = False
Timer20.Enabled = False
Timer36.Enabled = True
Timer40.Enabled = True
Text8.Text = Val(Text8.Text) + 100
Text7.Text = Val(Text8.Text) + 20
End If
End If
End Sub

Private Sub Timer68_Timer()
If Shape29.Visible = True Then
Command29.Enabled = True
Timer1.Interval = MMWheelX1.Value * 1000
Timer2.Interval = MMWheelX1.Value * 700
Timer3.Interval = MMWheelX1.Value * 600
Timer4.Interval = MMWheelX1.Value * 500
Timer5.Interval = MMWheelX1.Value * 400
Timer6.Interval = MMWheelX1.Value * 400
Timer7.Interval = MMWheelX1.Value * 400
Timer8.Interval = MMWheelX1.Value * 400
Timer9.Interval = MMWheelX1.Value * 900
Timer10.Interval = MMWheelX1.Value * 900
Timer11.Interval = MMWheelX1.Value * 900
Timer12.Interval = MMWheelX1.Value * 800
Timer13.Interval = MMWheelX1.Value * 500
Timer14.Interval = MMWheelX1.Value * 200
Timer15.Interval = MMWheelX1.Value * 200
Timer16.Interval = MMWheelX1.Value * 200
Timer17.Interval = MMWheelX1.Value * 400
Timer18.Interval = MMWheelX1.Value * 800
Timer19.Interval = MMWheelX1.Value * 900
End If
End Sub

Private Sub Timer69_Timer()
If MMSliderX1.Position = 3 Then
Command31.Enabled = False
Else
Command31.Enabled = True
End If
End Sub

Private Sub Timer7_Timer()
Timer7.Enabled = False
Shape13.Visible = False
Shape11.Visible = True
Timer8.Enabled = True
End Sub

Private Sub Timer70_Timer()
If Shape16.Visible = True Then
Command30.Enabled = True
Timer20.Interval = MMWheelX2.Value * 1000
Timer21.Interval = MMWheelX2.Value * 900
Timer22.Interval = MMWheelX2.Value * 800
Timer23.Interval = MMWheelX2.Value * 700
Timer24.Interval = MMWheelX2.Value * 600
Timer25.Interval = MMWheelX2.Value * 600
Timer26.Interval = MMWheelX2.Value * 600
Timer27.Interval = MMWheelX2.Value * 800
Timer28.Interval = MMWheelX2.Value * 700
Timer29.Interval = MMWheelX2.Value * 600
Timer30.Interval = MMWheelX2.Value * 500
Timer31.Interval = MMWheelX2.Value * 400
Timer32.Interval = MMWheelX2.Value * 200
Timer33.Interval = MMWheelX2.Value * 500
Timer34.Interval = MMWheelX2.Value * 700
End If
End Sub

Private Sub Timer8_Timer()
Timer8.Enabled = False
Shape11.Visible = False
Shape12.Visible = True
Timer9.Enabled = True
End Sub

Private Sub Timer9_Timer()
Timer9.Enabled = False
Shape17.Visible = True
Timer10.Enabled = True
Shape12.Visible = False
End Sub
