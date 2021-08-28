VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form3 
   BackColor       =   &H80000012&
   Caption         =   "Form3"
   ClientHeight    =   9930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13470
   LinkTopic       =   "Form3"
   ScaleHeight     =   9930
   ScaleWidth      =   13470
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command23 
      BackColor       =   &H00FF8080&
      Caption         =   "see last canidates result"
      Height          =   615
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   4320
      Width           =   1695
   End
   Begin VB.CommandButton Command22 
      BackColor       =   &H0080FF80&
      Caption         =   "go to previous page"
      Height          =   615
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton Command20 
      BackColor       =   &H00FF8080&
      Caption         =   "run test"
      Height          =   615
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   52
      ToolTipText     =   "click here only if you have saved your record on last run time"
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton Command19 
      Caption         =   "save result"
      Height          =   735
      Left            =   6480
      TabIndex        =   51
      Top             =   3720
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command18 
      BackColor       =   &H0080FF80&
      Caption         =   "end program"
      Height          =   615
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000015&
      Caption         =   "see old papers or create one"
      Height          =   1695
      Left            =   240
      TabIndex        =   46
      Top             =   3240
      Width           =   2175
      Begin VB.CommandButton Command11 
         BackColor       =   &H00FF8080&
         Caption         =   " break code"
         Height          =   375
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   360
         TabIndex        =   48
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton Command17 
         BackColor       =   &H00FF8080&
         Caption         =   "hide"
         Height          =   375
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   1080
         Width           =   1455
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3360
      Top             =   2640
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00FF8080&
      Caption         =   "refresh"
      Height          =   615
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   3600
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000015&
      Caption         =   "Question creater/modifier"
      ForeColor       =   &H8000000E&
      Height          =   4695
      Left            =   240
      TabIndex        =   25
      Top             =   5040
      Width           =   7215
      Begin VB.CommandButton Command21 
         BackColor       =   &H8000000D&
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "Haettenschweiler"
            Size            =   24
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   4680
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         DataField       =   "option1"
         DataSource      =   "Adodc1"
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   2160
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Check1"
         DataField       =   "option2"
         DataSource      =   "Adodc1"
         Height          =   195
         Left            =   120
         TabIndex        =   39
         Top             =   2400
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Check1"
         DataField       =   "option3"
         DataSource      =   "Adodc1"
         Height          =   195
         Left            =   120
         TabIndex        =   38
         Top             =   2640
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Check1"
         DataField       =   "option4"
         DataSource      =   "Adodc1"
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   2880
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF8080&
         Caption         =   "change question"
         Height          =   495
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H8000000D&
         Height          =   525
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   35
         Top             =   960
         Visible         =   0   'False
         Width           =   6975
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF8080&
         Caption         =   "ADD "
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   4200
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FF8080&
         Caption         =   "SAVE"
         Height          =   375
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   4200
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FF8080&
         Caption         =   "DELETE"
         Height          =   375
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   4200
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FF8080&
         Caption         =   "input 1st option name"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   3600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FF8080&
         Caption         =   "input 2nd option name"
         Height          =   495
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   3600
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FF8080&
         Caption         =   "input 3rd option name"
         Height          =   495
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   3600
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FF8080&
         Caption         =   "input 4th option name"
         Height          =   495
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   3600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H00FF8080&
         Caption         =   "clear checkboxes"
         Height          =   495
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton Command12 
         BackColor       =   &H00FF8080&
         Caption         =   "refresh"
         Height          =   375
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   4200
         Width           =   1095
      End
      Begin VB.Line Line5 
         X1              =   1680
         X2              =   1560
         Y1              =   480
         Y2              =   600
      End
      Begin VB.Line Line4 
         X1              =   1440
         X2              =   1560
         Y1              =   480
         Y2              =   600
      End
      Begin VB.Line Line3 
         X1              =   1560
         X2              =   1560
         Y1              =   240
         Y2              =   600
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Modifier"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   495
         Left            =   120
         TabIndex        =   55
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         DataField       =   "label1"
         DataSource      =   "Adodc1"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   360
         TabIndex        =   44
         Top             =   2160
         Visible         =   0   'False
         Width           =   6615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         DataField       =   "label2"
         DataSource      =   "Adodc1"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   360
         TabIndex        =   43
         Top             =   2400
         Visible         =   0   'False
         Width           =   6615
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         DataField       =   "label3"
         DataSource      =   "Adodc1"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   360
         TabIndex        =   42
         Top             =   2640
         Visible         =   0   'False
         Width           =   6615
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         DataField       =   "label4"
         DataSource      =   "Adodc1"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   360
         TabIndex        =   41
         Top             =   2880
         Visible         =   0   'False
         Width           =   6495
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      Caption         =   "Questions in storage"
      ForeColor       =   &H8000000E&
      Height          =   3135
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   7935
      Begin VB.CheckBox Check5 
         Caption         =   "Check1"
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   2640
         Width           =   255
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Check1"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   2400
         Width           =   255
      End
      Begin VB.CheckBox Check7 
         Caption         =   "Check1"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   2160
         Width           =   255
      End
      Begin VB.CheckBox Check8 
         Caption         =   "Check1"
         DataField       =   "question4"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   1920
         Width           =   255
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FF8080&
         Caption         =   "submit answer"
         Height          =   495
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   720
         Width           =   1335
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   2520
         Top             =   2640
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   495
         Left            =   360
         Top             =   240
         Width           =   2175
         _ExtentX        =   3836
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
         BackColor       =   -2147483642
         ForeColor       =   -2147483639
         Orientation     =   0
         Enabled         =   0
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Microsoft Visual Studio\VB98\hl logo\quiz.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Microsoft Visual Studio\VB98\hl logo\quiz.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "quiz"
         Caption         =   "question list"
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
      Begin VB.Label Text3 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   4680
         TabIndex        =   53
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404000&
         BackStyle       =   0  'Transparent
         Caption         =   "ASDASgggggggggggggggggggggggggggggggggggggggggggggggggggggggggggg"
         DataField       =   "question"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000014&
         Height          =   855
         Left            =   240
         TabIndex        =   24
         Top             =   840
         Width           =   5655
      End
      Begin VB.Label Label6 
         BackColor       =   &H00404000&
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         DataField       =   "label1"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000014&
         Height          =   255
         Left            =   600
         TabIndex        =   23
         Top             =   1920
         Width           =   7215
      End
      Begin VB.Label Label7 
         BackColor       =   &H00404000&
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         DataField       =   "label2"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000014&
         Height          =   255
         Left            =   600
         TabIndex        =   22
         Top             =   2160
         Width           =   7215
      End
      Begin VB.Label Label8 
         BackColor       =   &H00404000&
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         DataField       =   "label3"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000014&
         Height          =   255
         Left            =   600
         TabIndex        =   21
         Top             =   2400
         Width           =   7215
      End
      Begin VB.Label Label9 
         BackColor       =   &H00404000&
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         DataField       =   "label4"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000014&
         Height          =   255
         Left            =   600
         TabIndex        =   20
         Top             =   2640
         Width           =   7095
      End
      Begin VB.Label Label10 
         BackColor       =   &H00404000&
         BackStyle       =   0  'Transparent
         Caption         =   "current score"
         ForeColor       =   &H80000014&
         Height          =   495
         Left            =   3720
         TabIndex        =   19
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label11 
         BackColor       =   &H00404000&
         BackStyle       =   0  'Transparent
         Caption         =   "max marks"
         ForeColor       =   &H80000014&
         Height          =   375
         Left            =   5400
         TabIndex        =   18
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label12 
         BackColor       =   &H00404000&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H80000014&
         Height          =   375
         Left            =   6240
         TabIndex        =   17
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H00FF8080&
      Caption         =   "delete"
      Height          =   615
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1440
      Width           =   1695
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00FFC0C0&
      DataField       =   "totalmarks"
      DataSource      =   "Adodc2"
      Enabled         =   0   'False
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   9240
      TabIndex        =   9
      Text            =   "Text7"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00FFC0C0&
      DataField       =   "canidatescore"
      DataSource      =   "Adodc2"
      Enabled         =   0   'False
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   9240
      TabIndex        =   8
      Text            =   "Text6"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H00FF8080&
      Caption         =   "save"
      Height          =   615
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H00FF8080&
      Caption         =   "add"
      Height          =   615
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFC0C0&
      DataField       =   "canidaterol no"
      DataSource      =   "Adodc2"
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   9240
      TabIndex        =   1
      Text            =   "Text5"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFC0C0&
      DataField       =   "canidatename"
      DataSource      =   "Adodc2"
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   9240
      TabIndex        =   0
      Text            =   "Text4"
      Top             =   720
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   495
      Left            =   8640
      Top             =   3000
      Width           =   2295
      _ExtentX        =   4048
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
      BackColor       =   -2147483641
      ForeColor       =   -2147483639
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Microsoft Visual Studio\VB98\hl logo\quiz.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Microsoft Visual Studio\VB98\hl logo\quiz.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "score"
      Caption         =   "Adodc2"
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
   Begin VB.Label Label20 
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "click here to remove help"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   9360
      TabIndex        =   58
      Top             =   5520
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label19 
      BackColor       =   &H80000010&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "(before proceeding to step 1 click on ""add new"")"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   7680
      TabIndex        =   56
      Top             =   5880
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.Label Label16 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"Form3.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   3735
      Left            =   7680
      TabIndex        =   54
      Top             =   6120
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "max marks"
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   8160
      TabIndex        =   5
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "total score"
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   8160
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "roll no"
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   8160
      TabIndex        =   3
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "student name"
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   8160
      TabIndex        =   2
      Top             =   720
      Width           =   855
   End
   Begin VB.Line Line2 
      X1              =   -240
      X2              =   7800
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line1 
      X1              =   8040
      X2              =   8040
      Y1              =   0
      Y2              =   9960
   End
End
Attribute VB_Name = "form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Label1.Caption = Text1.Text
End Sub

Private Sub Command10_Click()
Check1.Value = Unchecked
Check2.Value = Unchecked
Check3.Value = Unchecked
Check4.Value = Unchecked
End Sub

Private Sub Command11_Click()
If Text2.Text = Text2.Tag Then
Command1.Visible = True
Command2.Visible = True
Command3.Visible = True
Command4.Visible = True
Command5.Visible = True
Command6.Visible = True
Command7.Visible = True
Command8.Visible = True
Command10.Visible = True
Check1.Visible = True
Check2.Visible = True
Check3.Visible = True
Check4.Visible = True
Text1.Visible = True
Label2.Visible = True
Label3.Visible = True
Label4.Visible = True
Label5.Visible = True
Text2.Text = Empty
Adodc1.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Else
MsgBox "incorrect code"
End If
End Sub

Private Sub Command12_Click()
Unload form3
Load form3
form3.Show
End Sub

Private Sub Command13_Click()
Unload form3
Load form3
form3.Show
End Sub

Private Sub Command14_Click()
On Error GoTo err:
Adodc2.Recordset.AddNew
Exit Sub
err:
End Sub

Private Sub Command15_Click()
On Error GoTo err
Adodc2.Recordset.Save
MsgBox "now start the program again and click on 'run test'"
End
Exit Sub
err:
End Sub

Private Sub Command16_Click()
On Error GoTo err:
Adodc2.Recordset.Delete
Adodc2.Recordset.MoveFirst
Exit Sub
err:
End Sub

Private Sub Command17_Click()
Command1.Visible = False
Command2.Visible = False
Command3.Visible = False
Command4.Visible = False
Command5.Visible = False
Command6.Visible = False
Command7.Visible = False
Command8.Visible = False
Command10.Visible = False
Check1.Visible = False
Check2.Visible = False
Check3.Visible = False
Check4.Visible = False
Text1.Visible = False
Label2.Visible = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Adodc1.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
End Sub

Private Sub Command18_Click()
End
End Sub

Private Sub Command19_Click()
Adodc2.Recordset.Save
Timer2.Enabled = True
End Sub

Private Sub Command2_Click()
On Error GoTo err:
Adodc1.Recordset.AddNew
Exit Sub
err:
End Sub

Private Sub Command20_Click()
Text5.Visible = False
Adodc2.Recordset.MoveLast
Dim userMsg As String
userMsg = InputBox("please match your identity", "roll no input", "Enter your stored roll no here", 500, 700)
If userMsg = Text5.Text Then
MsgBox "correct match proceeding with test"
Frame2.Visible = True
Command20.Enabled = False
Command14.Enabled = False
Command15.Enabled = False
Command16.Visible = False
Adodc2.Visible = False
Command13.Enabled = False
Else
MsgBox "wrong match!cannot proceed with test"
Adodc2.Recordset.MoveFirst
Text5.Visible = True
End If
End Sub

Private Sub Command21_Click()
Label16.Visible = True
Label19.Visible = True
Label20.Visible = True
End Sub

Private Sub Command22_Click()
form3.Hide
Form2.Show
End Sub

Private Sub Command23_Click()
form3.Hide
Form4.Show
Adodc2.Recordset.MoveLast
End Sub

Private Sub Command24_Click()
Print "hello"
End Sub

Private Sub Command3_Click()
On Error GoTo err
Adodc1.Recordset.Save
MsgBox "the program must restart before it can accept new data"
End
Exit Sub
err:
End Sub

Private Sub Command4_Click()
On Error GoTo err:
Adodc1.Recordset.Delete
Adodc1.Recordset.MoveFirst
Exit Sub
err:
End Sub

Private Sub Command5_Click()
Label2.Caption = Text1
Label6.Caption = Text1
End Sub

Private Sub Command6_Click()
Label3.Caption = Text1
Label7.Caption = Text1
End Sub

Private Sub Command7_Click()
Label4.Caption = Text1
Label8.Caption = Text1
End Sub

Private Sub Command8_Click()
Label5.Caption = Text1
Label9.Caption = Text1
End Sub

Private Sub Command9_Click()
Label12.Caption = Val(Label12) + 1
If Check8.Value = Check1.Value And Check2.Value = Unchecked And Check3.Value = Unchecked And Check4.Value = Unchecked Then
MsgBox "correct"
Text3.Caption = Val(Text3.Caption) + 1
Else
If Check7.Value = Check2.Value And Check1.Value = Unchecked And Check3.Value = Unchecked And Check4.Value = Unchecked Then
MsgBox "correct"
Text3.Caption = Val(Text3.Caption) + 1
Else
If Check6.Value = Check3.Value And Check1.Value = Unchecked And Check2.Value = Unchecked And Check4.Value = Unchecked Then
MsgBox "correct"
Text3.Caption = Val(Text3.Caption) + 1
Else
If Check5.Value = Check4.Value And Check1.Value = Unchecked And Check2.Value = Unchecked And Check3.Value = Unchecked Then
MsgBox "correct"
Text3.Caption = Val(Text3.Caption) + 1
Else
MsgBox "incorrect"
End If
End If
End If
End If
Adodc1.Recordset.MoveNext
Check8.Value = 0
Check7.Value = 0
Check6.Value = 0
Check5.Value = 0
If Adodc1.Recordset.EOF Then
Frame2.Enabled = False
MsgBox "quiz finished"
End If
End Sub

Private Sub Label20_Click()
Label16.Visible = False
Label19.Visible = False
Label20.Visible = False
End Sub

Private Sub Timer1_Timer()
If Adodc1.Recordset.EOF Then
Timer1.Enabled = False
Text6.Text = Text3.Caption
Text7.Text = Label12.Caption
Frame2.Visible = True
Command14.Enabled = True
Command15.Enabled = True
Command16.Visible = True
Command13.Enabled = True
Adodc2.Recordset.Save
Timer2.Enabled = True
End If
End Sub

Private Sub Timer2_Timer()
form3.Hide
Form4.Show
Timer2.Enabled = False
End Sub
