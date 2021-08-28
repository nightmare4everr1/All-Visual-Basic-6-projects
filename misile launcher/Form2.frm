VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   9525
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13935
   LinkTopic       =   "Form2"
   ScaleHeight     =   9525
   ScaleWidth      =   13935
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   6360
      Top             =   3480
      Width           =   1815
      _ExtentX        =   3201
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
      Connect         =   $"Form2.frx":0000
      OLEDBString     =   $"Form2.frx":0089
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
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Hide
Form1.Show
str1 = Label1.Caption
    strsearch = "name like '" & str1 & "'"

    Form1.Adodc1.Recordset.MoveFirst
    Form1.Adodc1.Recordset.Filter = (strsearch)
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

Private Sub Form_Load()

End Sub
