VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10815
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   ScaleHeight     =   10815
   ScaleWidth      =   11400
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command7 
      Caption         =   "goto form2"
      Height          =   615
      Left            =   9480
      TabIndex        =   29
      Top             =   7320
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   9240
      Top             =   2520
   End
   Begin VB.Frame Frame4 
      Caption         =   "for adminisrator"
      Height          =   2055
      Left            =   9000
      TabIndex        =   25
      Top             =   5160
      Width           =   2295
      Begin VB.CommandButton Command6 
         Caption         =   "disable break"
         Height          =   495
         Left            =   360
         TabIndex        =   28
         Top             =   1320
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   360
         PasswordChar    =   "*"
         TabIndex        =   27
         Tag             =   "gordon"
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command5 
         Caption         =   "break code"
         Height          =   495
         Left            =   360
         TabIndex        =   26
         Top             =   720
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "refresh"
      Height          =   495
      Left            =   4200
      TabIndex        =   22
      ToolTipText     =   "click on this whenever you delete a record or if you find any programme errors"
      Top             =   10200
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "click to start voting"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2280
      TabIndex        =   21
      ToolTipText     =   "clicking on this control will add a new space in the database (caution:all previous unsaved data will be lost)"
      Top             =   1800
      Width           =   4815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "delete"
      Height          =   495
      Left            =   6480
      TabIndex        =   20
      ToolTipText     =   "thiswill delete the current showing record permanently"
      Top             =   10200
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000A&
      Caption         =   "save"
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
      Left            =   1320
      MaskColor       =   &H80000000&
      TabIndex        =   19
      ToolTipText     =   "this wil cause the inputted values to be saved"
      Top             =   9240
      Width           =   6975
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   840
      Top             =   10200
      Width           =   3255
      _ExtentX        =   5741
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
      Connect         =   $"Form1.frx":0000
      OLEDBString     =   $"Form1.frx":0099
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "poll"
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
   Begin VB.Frame Frame3 
      Caption         =   "choose your favourite"
      Enabled         =   0   'False
      Height          =   3135
      Left            =   480
      TabIndex        =   14
      Top             =   6000
      Width           =   8175
      Begin VB.ComboBox Combo3 
         DataField       =   "your vote goes for"
         DataSource      =   "Adodc1"
         Height          =   315
         ItemData        =   "Form1.frx":0132
         Left            =   720
         List            =   "Form1.frx":013F
         TabIndex        =   16
         ToolTipText     =   $"Form1.frx":0177
         Top             =   1800
         Width           =   6855
      End
      Begin VB.Label Label9 
         Caption         =   "to whom wil you give your popularity vote?"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1215
         Left            =   1080
         TabIndex        =   15
         Top             =   360
         Width           =   6495
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "personal details"
      Enabled         =   0   'False
      Height          =   3135
      Left            =   480
      TabIndex        =   2
      Top             =   2760
      Width           =   8175
      Begin VB.ComboBox Combo6 
         DataField       =   "province"
         DataSource      =   "Adodc1"
         Height          =   315
         ItemData        =   "Form1.frx":0228
         Left            =   5640
         List            =   "Form1.frx":023B
         TabIndex        =   24
         ToolTipText     =   $"Form1.frx":026A
         Top             =   1080
         Width           =   2175
      End
      Begin VB.ComboBox Combo5 
         DataField       =   "sex"
         DataSource      =   "Adodc1"
         Height          =   315
         ItemData        =   "Form1.frx":031B
         Left            =   1800
         List            =   "Form1.frx":0325
         TabIndex        =   23
         ToolTipText     =   $"Form1.frx":0337
         Top             =   1800
         Width           =   2175
      End
      Begin VB.ComboBox Combo4 
         DataField       =   "do you know politics fairly"
         DataSource      =   "Adodc1"
         Height          =   315
         ItemData        =   "Form1.frx":03E8
         Left            =   3000
         List            =   "Form1.frx":03F5
         TabIndex        =   18
         ToolTipText     =   $"Form1.frx":0409
         Top             =   2400
         Width           =   2295
      End
      Begin VB.ComboBox Combo2 
         DataField       =   "rular or urban"
         DataSource      =   "Adodc1"
         Height          =   315
         ItemData        =   "Form1.frx":04BA
         Left            =   5640
         List            =   "Form1.frx":04C4
         TabIndex        =   13
         ToolTipText     =   $"Form1.frx":04D6
         Top             =   1800
         Width           =   2175
      End
      Begin VB.ComboBox Combo1 
         DataField       =   "nationality"
         DataSource      =   "Adodc1"
         Height          =   315
         ItemData        =   "Form1.frx":0587
         Left            =   5640
         List            =   "Form1.frx":0591
         TabIndex        =   11
         ToolTipText     =   $"Form1.frx":05AC
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         DataField       =   "name"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   1800
         TabIndex        =   4
         Top             =   480
         Width           =   2175
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
      Begin VB.Label Label10 
         Caption         =   "do you have a fair idea of political situations?"
         Height          =   495
         Left            =   360
         TabIndex        =   17
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   "rular or urban"
         Height          =   495
         Left            =   4080
         TabIndex        =   12
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "nationality"
         Height          =   375
         Left            =   4080
         TabIndex        =   10
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "province"
         Height          =   375
         Left            =   4080
         TabIndex        =   9
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "your name"
         Height          =   375
         Left            =   480
         TabIndex        =   8
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "gender"
         Height          =   375
         Left            =   480
         TabIndex        =   7
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label4 
         Height          =   375
         Left            =   4080
         TabIndex        =   6
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "age"
         Height          =   375
         Left            =   480
         TabIndex        =   5
         Top             =   1200
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   8295
      Begin VB.Label Label1 
         Caption         =   "the gilani poll"
         BeginProperty Font 
            Name            =   "HL2cross"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   615
         Left            =   2400
         TabIndex        =   1
         Top             =   240
         Width           =   3855
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo err
Adodc1.Recordset.Save
Frame2.Enabled = False
Frame3.Enabled = False
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
On Error GoTo err
Frame2.Enabled = True
Frame3.Enabled = True
Adodc1.Recordset.AddNew
Exit Sub
err:
End Sub

Private Sub Command4_Click()
Unload Form1
Load Form1
Form1.Show
End Sub

Private Sub Command5_Click()
If Text3.Text = Text3.Tag Then
Frame2.Enabled = True
Frame3.Enabled = True
Command6.Visible = True
Command3.Enabled = False
Command4.Enabled = False
End If
End Sub

Private Sub Command6_Click()
Text3.Text = Empty
Frame2.Enabled = False
Frame3.Enabled = False
Command3.Enabled = True
Command4.Enabled = True
Command6.Visible = False
End Sub

Private Sub Command7_Click()
Form1.Hide
Form2.Show
End Sub

Private Sub Form_Load()
MsgBox "CAUTION:DO NOT ATTEMPT TO WRITE EXCEPT WHATS WRITTEN IN THE PULL DOWN INTERFACE.DOING SO WILL RESULT IN TERMINATION OF YOUR ENTRY"
End Sub


Private Sub Timer1_Timer()
If Combo1.Text = "pakistani" Then
Combo6.Enabled = True
Else
Combo6.Enabled = False
End If
End Sub
