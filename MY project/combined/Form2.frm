VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{A40B4820-D5FD-11D1-8818-C199198E9702}#1.8#0"; "MMTOOLSX.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   9870
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13455
   LinkTopic       =   "Form2"
   ScaleHeight     =   9870
   ScaleWidth      =   13455
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MMToolsX.MMGaugeX MMGaugeX1 
      Height          =   1455
      Left            =   240
      TabIndex        =   17
      Top             =   7200
      Width           =   15135
      ForeColor       =   4210688
      MaxValue        =   1000
      Progress        =   1000
      _Handle         =   1
   End
   Begin VB.CommandButton Command6 
      Height          =   735
      Left            =   3720
      TabIndex        =   30
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Height          =   615
      Left            =   11640
      TabIndex        =   16
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton Command10 
      Caption         =   "search"
      Height          =   615
      Left            =   9840
      TabIndex        =   15
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton Command8 
      Caption         =   "enable"
      Height          =   495
      Left            =   4200
      TabIndex        =   14
      Top             =   9360
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "disable"
      Height          =   495
      Left            =   4200
      TabIndex        =   13
      Top             =   8760
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   0
      Picture         =   "Form2.frx":0000
      ScaleHeight     =   675
      ScaleWidth      =   4035
      TabIndex        =   12
      Top             =   3240
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
      Left            =   7800
      Top             =   3360
   End
   Begin VB.Timer Timer9 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9600
      Top             =   8640
   End
   Begin VB.TextBox Text2 
      DataField       =   "lives left"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   8040
      TabIndex        =   11
      Top             =   9720
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      DataField       =   "score"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   8040
      TabIndex        =   10
      Top             =   9120
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "delete"
      Height          =   375
      Left            =   6960
      TabIndex        =   9
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      DataField       =   "name"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   615
      Left            =   6960
      TabIndex        =   8
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "add new"
      Height          =   375
      Left            =   6960
      TabIndex        =   7
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "refresh/restart"
      Height          =   375
      Left            =   6000
      TabIndex        =   6
      Top             =   0
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "save"
      Height          =   735
      Left            =   9960
      TabIndex        =   5
      Top             =   9120
      Width           =   1695
   End
   Begin VB.Timer Timer8 
      Interval        =   1
      Left            =   4680
      Top             =   6120
   End
   Begin VB.Timer Timer7 
      Interval        =   1
      Left            =   4680
      Top             =   5640
   End
   Begin VB.Timer Timer6 
      Interval        =   1
      Left            =   4680
      Top             =   5160
   End
   Begin VB.CommandButton Command1 
      Caption         =   "start"
      Enabled         =   0   'False
      Height          =   735
      Left            =   6000
      TabIndex        =   4
      Top             =   480
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
      Left            =   1080
      TabIndex        =   3
      Top             =   1440
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
      Left            =   1080
      TabIndex        =   2
      Top             =   600
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
      Left            =   1080
      TabIndex        =   1
      Top             =   3120
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
      Left            =   1080
      TabIndex        =   0
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer5 
      Interval        =   1
      Left            =   4680
      Top             =   4680
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   2400
      Top             =   3000
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   2400
      Top             =   2280
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   2400
      Top             =   1440
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   2400
      Top             =   840
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "Form2.frx":105F
      Height          =   2895
      Left            =   9960
      TabIndex        =   18
      Top             =   1800
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
   Begin VB.Label Label12 
      Caption         =   "enter your name here"
      Height          =   495
      Left            =   4800
      TabIndex        =   29
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label10 
      Caption         =   "lives left"
      Height          =   495
      Left            =   6720
      TabIndex        =   28
      Top             =   9720
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "score"
      Height          =   375
      Left            =   6720
      TabIndex        =   27
      Top             =   9120
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
      Left            =   5160
      TabIndex        =   26
      Top             =   6120
      Width           =   7815
   End
   Begin VB.Label Label7 
      Caption         =   "surface"
      Height          =   495
      Left            =   5640
      TabIndex        =   25
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Line Line1 
      X1              =   480
      X2              =   5520
      Y1              =   3840
      Y2              =   3840
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
      Left            =   360
      TabIndex        =   24
      Top             =   6000
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
      Left            =   600
      TabIndex        =   23
      Top             =   4920
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
      Left            =   840
      TabIndex        =   22
      Top             =   3960
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
      Left            =   1800
      TabIndex        =   21
      Top             =   6000
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
      Left            =   2040
      TabIndex        =   20
      Top             =   5040
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
      Left            =   2040
      TabIndex        =   19
      Top             =   3960
      Width           =   1575
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command6_Click()
Form2.Hide
Form1.Show
End Sub
