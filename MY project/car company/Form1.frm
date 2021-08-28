VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13875
   LinkTopic       =   "Form1"
   ScaleHeight     =   11115
   ScaleWidth      =   13875
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command6 
      Caption         =   "my style delete"
      Height          =   735
      Left            =   5040
      TabIndex        =   69
      Top             =   9360
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "clear"
      Height          =   735
      Left            =   9120
      TabIndex        =   68
      Top             =   3360
      Width           =   1815
   End
   Begin VB.TextBox Text18 
      DataField       =   "cost"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   2280
      TabIndex        =   67
      Top             =   330
      Width           =   2295
   End
   Begin VB.CheckBox Check8 
      Caption         =   "Check1"
      DataField       =   "sunroof"
      DataSource      =   "Adodc1"
      Height          =   195
      Left            =   6480
      TabIndex        =   65
      Top             =   6000
      Width           =   255
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Check2"
      DataField       =   "sunroofn"
      DataSource      =   "Adodc1"
      Height          =   195
      Left            =   7560
      TabIndex        =   64
      Top             =   6000
      Width           =   255
   End
   Begin VB.TextBox Text17 
      DataField       =   "consumptiond"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   9000
      TabIndex        =   62
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox Text16 
      DataField       =   "consumptionc"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   7920
      TabIndex        =   61
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox Text15 
      DataField       =   "consumptionb"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   6840
      TabIndex        =   60
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox Text14 
      DataField       =   "consumptiona"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   5880
      TabIndex        =   59
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "refresh"
      Height          =   495
      Left            =   2760
      TabIndex        =   57
      Top             =   9720
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "delete"
      Height          =   495
      Left            =   240
      TabIndex        =   56
      Top             =   9720
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "save"
      Height          =   615
      Left            =   2760
      TabIndex        =   55
      Top             =   10320
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "add new"
      Height          =   615
      Left            =   240
      TabIndex        =   54
      Top             =   10320
      Width           =   2295
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   5280
      Top             =   10440
      Width           =   2775
      _ExtentX        =   4895
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
      OLEDBString     =   $"Form1.frx":0096
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "carcompany"
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
   Begin VB.CheckBox Check4 
      Caption         =   "Check1"
      DataField       =   "fuelc"
      DataSource      =   "Adodc1"
      Height          =   195
      Left            =   8160
      TabIndex        =   47
      Top             =   720
      Width           =   255
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Check2"
      DataField       =   "fueld"
      DataSource      =   "Adodc1"
      Height          =   195
      Left            =   9240
      TabIndex        =   46
      Top             =   720
      Width           =   255
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check1"
      DataField       =   "fuela"
      DataSource      =   "Adodc1"
      Height          =   195
      Left            =   6000
      TabIndex        =   45
      Top             =   720
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check2"
      DataField       =   "fuelb"
      DataSource      =   "Adodc1"
      Height          =   195
      Left            =   7080
      TabIndex        =   44
      Top             =   720
      Width           =   255
   End
   Begin VB.TextBox Text13 
      DataField       =   "listofcolurs"
      DataSource      =   "Adodc1"
      Height          =   615
      Left            =   4920
      TabIndex        =   43
      Top             =   8280
      Width           =   3135
   End
   Begin VB.CheckBox Check6 
      Caption         =   "Check1"
      DataField       =   "abs"
      DataSource      =   "Adodc1"
      Height          =   195
      Left            =   6480
      TabIndex        =   35
      Top             =   4680
      Width           =   255
   End
   Begin VB.CheckBox Check7 
      Caption         =   "Check2"
      DataField       =   "absn"
      DataSource      =   "Adodc1"
      Height          =   195
      Left            =   7560
      TabIndex        =   34
      Top             =   4680
      Width           =   255
   End
   Begin VB.CheckBox Check11 
      Caption         =   "Check1"
      DataField       =   "autogear"
      DataSource      =   "Adodc1"
      Height          =   195
      Left            =   6480
      TabIndex        =   33
      Top             =   3840
      Width           =   255
   End
   Begin VB.CheckBox Check12 
      Caption         =   "Check2"
      DataField       =   "autogearn"
      DataSource      =   "Adodc1"
      Height          =   195
      Left            =   7560
      TabIndex        =   32
      Top             =   3840
      Width           =   255
   End
   Begin VB.CheckBox Check16 
      Caption         =   "Check1"
      DataField       =   "motorwindow"
      DataSource      =   "Adodc1"
      Height          =   195
      Left            =   6480
      TabIndex        =   31
      Top             =   3120
      Width           =   255
   End
   Begin VB.CheckBox Check17 
      Caption         =   "Check2"
      DataField       =   "motorwindown"
      DataSource      =   "Adodc1"
      Height          =   195
      Left            =   7560
      TabIndex        =   30
      Top             =   3120
      Width           =   255
   End
   Begin VB.CheckBox Check21 
      Caption         =   "Check1"
      DataField       =   "autobrake"
      DataSource      =   "Adodc1"
      Height          =   195
      Left            =   6480
      TabIndex        =   29
      Top             =   2400
      Width           =   255
   End
   Begin VB.CheckBox Check22 
      Caption         =   "Check2"
      DataField       =   "autobraken"
      DataSource      =   "Adodc1"
      Height          =   195
      Left            =   7560
      TabIndex        =   28
      Top             =   2400
      Width           =   255
   End
   Begin VB.CheckBox Check26 
      Caption         =   "Check1"
      DataField       =   "safetybag"
      DataSource      =   "Adodc1"
      Height          =   195
      Left            =   6480
      TabIndex        =   27
      Top             =   5400
      Width           =   255
   End
   Begin VB.CheckBox Check27 
      Caption         =   "Check2"
      DataField       =   "safetybagn"
      DataSource      =   "Adodc1"
      Height          =   195
      Left            =   7560
      TabIndex        =   26
      Top             =   5400
      Width           =   255
   End
   Begin VB.TextBox Text12 
      DataField       =   "wheeldia"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2400
      TabIndex        =   24
      Top             =   9000
      Width           =   1695
   End
   Begin VB.TextBox Text11 
      DataField       =   "brakesize"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2400
      TabIndex        =   20
      Top             =   8400
      Width           =   1695
   End
   Begin VB.TextBox Text10 
      DataField       =   "rideheight"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   3000
      TabIndex        =   19
      Top             =   7200
      Width           =   1575
   End
   Begin VB.TextBox Text9 
      DataField       =   "acceleration"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   2760
      TabIndex        =   16
      Top             =   3240
      Width           =   615
   End
   Begin VB.TextBox Text8 
      DataField       =   "mass"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2400
      TabIndex        =   14
      Top             =   7800
      Width           =   1695
   End
   Begin VB.TextBox Text7 
      DataField       =   "frontrear"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   1080
      TabIndex        =   12
      Top             =   7200
      Width           =   1695
   End
   Begin VB.TextBox Text6 
      DataField       =   "enginedetails"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   1680
      TabIndex        =   10
      Top             =   5400
      Width           =   2535
   End
   Begin VB.TextBox Text5 
      DataField       =   "horsepower"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2520
      TabIndex        =   8
      Top             =   4800
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      DataField       =   "speed"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2280
      TabIndex        =   6
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      DataField       =   "model"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      DataField       =   "company"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      DataField       =   "car"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label16 
      Caption         =   "cost(pak rs)"
      Height          =   375
      Left            =   1200
      TabIndex        =   66
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label30 
      Caption         =   "sun roof"
      Height          =   375
      Left            =   5280
      TabIndex        =   63
      Top             =   6000
      Width           =   975
   End
   Begin VB.Label Label29 
      Caption         =   "fuel consumption(/km)"
      Height          =   375
      Left            =   4200
      TabIndex        =   58
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label28 
      Caption         =   "high octane"
      Height          =   375
      Left            =   8880
      TabIndex        =   53
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label27 
      Caption         =   "petrol"
      Height          =   375
      Left            =   7920
      TabIndex        =   52
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label26 
      Caption         =   "cng"
      Height          =   375
      Left            =   6960
      TabIndex        =   51
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label25 
      Caption         =   "diesel"
      Height          =   375
      Left            =   5880
      TabIndex        =   50
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label24 
      Caption         =   "no"
      Height          =   375
      Left            =   7560
      TabIndex        =   49
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label23 
      Caption         =   "yes"
      Height          =   495
      Left            =   6480
      TabIndex        =   48
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label22 
      Caption         =   "list of colours available"
      Height          =   375
      Left            =   5400
      TabIndex        =   42
      Top             =   7920
      Width           =   1935
   End
   Begin VB.Label Label21 
      Caption         =   "safety wheel bag"
      Height          =   495
      Left            =   5280
      TabIndex        =   41
      Top             =   5280
      Width           =   855
   End
   Begin VB.Label Label20 
      Caption         =   "electricel motor powerer rool up windows"
      Height          =   615
      Left            =   4920
      TabIndex        =   40
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label19 
      Caption         =   "automatic gear system"
      Height          =   495
      Left            =   5280
      TabIndex        =   39
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label Label18 
      Caption         =   "anti-lock brakin system"
      Height          =   495
      Left            =   5040
      TabIndex        =   38
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label Label17 
      Caption         =   "fuel intake"
      Height          =   495
      Left            =   4800
      TabIndex        =   37
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label15 
      Caption         =   "automatic braking system"
      Height          =   615
      Left            =   5280
      TabIndex        =   36
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label14 
      Caption         =   "wheel diameter(cm)"
      Height          =   495
      Left            =   960
      TabIndex        =   25
      Top             =   9000
      Width           =   1815
   End
   Begin VB.Label Label13 
      Caption         =   "brake axle size(cm)"
      Height          =   495
      Left            =   1080
      TabIndex        =   23
      Top             =   8400
      Width           =   1215
   End
   Begin VB.Label Label12 
      Caption         =   "ride height"
      Height          =   375
      Left            =   3120
      TabIndex        =   22
      Top             =   6720
      Width           =   1455
   End
   Begin VB.Label Label11 
      Caption         =   "front to rear"
      Height          =   255
      Left            =   1200
      TabIndex        =   21
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Label Label10 
      Caption         =   "seconds"
      Height          =   375
      Left            =   3360
      TabIndex        =   18
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "acceleration:1-60 in"
      Height          =   495
      Left            =   1200
      TabIndex        =   17
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label8 
      Caption         =   "mass unloaded"
      Height          =   495
      Left            =   1080
      TabIndex        =   15
      Top             =   7800
      Width           =   1815
   End
   Begin VB.Label Label7 
      Caption         =   "dimensions"
      Height          =   495
      Left            =   2040
      TabIndex        =   13
      Top             =   6240
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "engine details"
      Height          =   495
      Left            =   480
      TabIndex        =   11
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "horse power"
      Height          =   495
      Left            =   1200
      TabIndex        =   9
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "speed"
      Height          =   495
      Left            =   1200
      TabIndex        =   7
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "model"
      Height          =   495
      Left            =   1200
      TabIndex        =   5
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "company"
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "car name"
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   720
      Width           =   1815
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
Check1.Value = Unchecked
Check2.Value = Unchecked
Check3.Value = Unchecked
Check4.Value = Unchecked
Check5.Value = Unchecked
Check6.Value = Unchecked
Check7.Value = Unchecked
Check8.Value = Unchecked
Check11.Value = Unchecked
Check12.Value = Unchecked
Check16.Value = Unchecked
Check17.Value = Unchecked
Check21.Value = Unchecked
Check22.Value = Unchecked
Check26.Value = Unchecked
Check27.Value = Unchecked
End Sub

Private Sub Command6_Click()
Check1.Value = Unchecked
Check2.Value = Unchecked
Check3.Value = Unchecked
Check4.Value = Unchecked
Check5.Value = Unchecked
Check6.Value = Unchecked
Check7.Value = Unchecked
Check8.Value = Unchecked
Check11.Value = Unchecked
Check12.Value = Unchecked
Check16.Value = Unchecked
Check17.Value = Unchecked
Check21.Value = Unchecked
Check22.Value = Unchecked
Check26.Value = Unchecked
Check27.Value = Unchecked
Text1.Text = Empty
Text2.Text = Empty
Text3.Text = Empty
Text4.Text = Empty
Text5.Text = Empty
Text6.Text = Empty
Text7.Text = Empty
Text7.Text = Empty
Text8.Text = Empty
Text9.Text = Empty
Text10.Text = Empty
Text11.Text = Empty
Text12.Text = Empty
Text13.Text = Empty
Text14.Text = Empty
Text15.Text = Empty
Text16.Text = Empty
Text17.Text = Empty
Text18.Text = Empty
End Sub
