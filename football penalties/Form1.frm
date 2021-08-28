VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   BackColor       =   &H00808000&
   Caption         =   "Penalty Simulator"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   495
      Left            =   1440
      TabIndex        =   29
      Top             =   7440
      Width           =   1335
   End
   Begin VB.PictureBox MediaPlayer1 
      Height          =   375
      Left            =   8760
      ScaleHeight     =   315
      ScaleWidth      =   2235
      TabIndex        =   27
      Top             =   5280
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Timer Timer18 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5520
      Top             =   9960
   End
   Begin VB.Timer Timer17 
      Interval        =   1
      Left            =   5280
      Top             =   7920
   End
   Begin VB.Timer Timer16 
      Interval        =   1
      Left            =   4800
      Top             =   7920
   End
   Begin VB.Timer Timer15 
      Interval        =   1
      Left            =   4800
      Top             =   7440
   End
   Begin VB.Timer Timer14 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5760
      Top             =   7440
   End
   Begin VB.Timer Timer13 
      Interval        =   1
      Left            =   5280
      Top             =   7440
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   7440
      TabIndex        =   23
      Text            =   "Text5"
      Top             =   5280
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   7440
      TabIndex        =   20
      Text            =   "0"
      Top             =   4920
      Width           =   735
   End
   Begin VB.Timer Timer12 
      Interval        =   1
      Left            =   7920
      Top             =   5880
   End
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5760
      Top             =   2880
   End
   Begin VB.Timer Timer11 
      Interval        =   1
      Left            =   1320
      Top             =   840
   End
   Begin VB.Timer Timer10 
      Interval        =   1
      Left            =   600
      Top             =   6120
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FF8080&
      Caption         =   "fool the keeper"
      Height          =   735
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6120
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H000040C0&
      Caption         =   "reset full"
      Height          =   375
      Left            =   240
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H8000000D&
      Caption         =   "help"
      Height          =   495
      Left            =   240
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Timer Timer9 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   6720
   End
   Begin VB.Timer Timer8 
      Interval        =   1
      Left            =   4440
      Top             =   960
   End
   Begin VB.Timer Timer4 
      Interval        =   1
      Left            =   6960
      Top             =   6360
   End
   Begin VB.Timer Timer6 
      Interval        =   1
      Left            =   3240
      Top             =   480
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H000000C0&
      Caption         =   "console"
      Height          =   375
      Left            =   240
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Timer Timer5 
      Interval        =   1
      Left            =   7680
      Top             =   3720
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6960
      Top             =   5880
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   7440
      Top             =   6360
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   4320
      TabIndex        =   5
      Text            =   "100"
      Top             =   6360
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   4320
      TabIndex        =   4
      Text            =   "100"
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00008000&
      Caption         =   "reset"
      Height          =   375
      Left            =   240
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Text            =   "0"
      Top             =   5400
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   7440
      Top             =   5880
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "shoot"
      Height          =   735
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5280
      Width           =   1695
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   495
      Left            =   9240
      TabIndex        =   28
      Top             =   4800
      Visible         =   0   'False
      Width           =   1815
      URL             =   "C:\RECORDER\favela rock afrobots.mp3"
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
      _cx             =   3201
      _cy             =   873
   End
   Begin VB.Label Label16 
      Caption         =   "0"
      Height          =   375
      Left            =   4200
      TabIndex        =   26
      Top             =   9960
      Width           =   855
   End
   Begin VB.Line Line12 
      X1              =   12480
      X2              =   12480
      Y1              =   7320
      Y2              =   9480
   End
   Begin VB.Line Line13 
      X1              =   13440
      X2              =   7080
      Y1              =   9480
      Y2              =   9480
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   7320
      Shape           =   3  'Circle
      Top             =   9120
      Width           =   495
   End
   Begin VB.Line Line14 
      X1              =   12480
      X2              =   13440
      Y1              =   7320
      Y2              =   7920
   End
   Begin VB.Line Line15 
      X1              =   13440
      X2              =   13440
      Y1              =   7920
      Y2              =   9480
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00404000&
      FillStyle       =   0  'Solid
      Height          =   1335
      Left            =   12240
      Shape           =   4  'Rounded Rectangle
      Top             =   8160
      Width           =   255
   End
   Begin VB.Line Line16 
      X1              =   7080
      X2              =   7320
      Y1              =   9480
      Y2              =   9480
   End
   Begin VB.Shape Shape7 
      BackStyle       =   1  'Opaque
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   3015
      Left            =   6840
      Top             =   6840
      Width           =   7215
   End
   Begin VB.Label Label15 
      BackColor       =   &H0000FFFF&
      Caption         =   "Side Angle View"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   615
      Left            =   8280
      TabIndex        =   25
      Top             =   10080
      Width           =   4095
   End
   Begin VB.Label Label14 
      Caption         =   "1"
      Height          =   375
      Left            =   10080
      TabIndex        =   24
      Top             =   1560
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "wind direction"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   6240
      TabIndex        =   22
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "wind speed"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   6240
      TabIndex        =   21
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Left            =   360
      TabIndex        =   19
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Label10"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   4800
      TabIndex        =   18
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "keeper"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4920
      TabIndex        =   17
      Top             =   2400
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000006&
      BorderWidth     =   2
      FillColor       =   &H00FF0000&
      Height          =   855
      Left            =   4320
      Shape           =   3  'Circle
      Top             =   3000
      Width           =   855
   End
   Begin VB.Line Line3 
      BorderWidth     =   4
      X1              =   2400
      X2              =   7200
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00404000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      FillColor       =   &H008080FF&
      Height          =   1335
      Left            =   4920
      Shape           =   4  'Rounded Rectangle
      Top             =   2400
      Width           =   615
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000007&
      BorderWidth     =   4
      DrawMode        =   1  'Blackness
      X1              =   7200
      X2              =   7200
      Y1              =   1440
      Y2              =   3720
   End
   Begin VB.Line Line1 
      BorderWidth     =   4
      X1              =   2400
      X2              =   2400
      Y1              =   1440
      Y2              =   3720
   End
   Begin VB.Line Line4 
      X1              =   2400
      X2              =   3360
      Y1              =   1440
      Y2              =   1200
   End
   Begin VB.Line Line5 
      X1              =   3360
      X2              =   8040
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line7 
      X1              =   8040
      X2              =   8040
      Y1              =   1200
      Y2              =   3360
   End
   Begin VB.Line Line9 
      X1              =   3360
      X2              =   3360
      Y1              =   1200
      Y2              =   3360
   End
   Begin VB.Line Line11 
      X1              =   3360
      X2              =   8040
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Shape Shape4 
      BorderStyle     =   6  'Inside Solid
      Height          =   2175
      Left            =   3360
      Top             =   1200
      Width           =   4695
   End
   Begin VB.Line Line10 
      X1              =   2400
      X2              =   3360
      Y1              =   3720
      Y2              =   3360
   End
   Begin VB.Line Line8 
      X1              =   7200
      X2              =   8040
      Y1              =   3720
      Y2              =   3360
   End
   Begin VB.Line Line6 
      X1              =   8040
      X2              =   7200
      Y1              =   1200
      Y2              =   1440
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      Height          =   4215
      Left            =   1560
      Top             =   720
      Width           =   7335
   End
   Begin VB.Label Label8 
      Caption         =   "20"
      Height          =   375
      Left            =   5400
      TabIndex        =   16
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "40"
      Height          =   255
      Left            =   5400
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "10"
      Height          =   255
      Left            =   9960
      TabIndex        =   10
      Top             =   2160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "10000"
      Height          =   375
      Left            =   7200
      TabIndex        =   9
      Top             =   4200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "1"
      Height          =   255
      Left            =   9840
      TabIndex        =   8
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "upward direction"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   4320
      TabIndex        =   7
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "horizontal direction"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "power"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3000
      TabIndex        =   2
      Top             =   5040
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Timer1.Enabled = True
Timer1.Interval = 150 - Val(Text1.Text)
Timer2.Enabled = True
Timer10.Enabled = False
If Label11.Caption = "1" Then
Timer7.Enabled = True
Timer18.Enabled = True
End If
End Sub

Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer9.Enabled = True
End Sub

Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer9.Enabled = False
End Sub

Private Sub Command2_Click()
For X = 1 To 10
X = Rnd * 30
Next X
For Y = 1 To 10
Y = Rnd * 50
Next Y
If Y >= 25 Then
X = -X
Else
X = X
End If
Text4.Text = X
Timer1.Enabled = False
Timer3.Enabled = False
Timer2.Enabled = False
Timer5.Enabled = True
Label6.Caption = 10
Label4.Caption = 1
Timer8.Enabled = True
Shape1.Height = 855
Shape1.Left = 4320
Shape1.Top = 3000
Text1.Text = 1
Timer7.Enabled = False
If Label11.Caption = "1" Then
Timer10.Enabled = True
End If
Timer13.Enabled = True
Timer14.Enabled = False
Timer18.Enabled = False
End Sub

Private Sub Command3_Click()
Dim X As String
X = InputBox("console", "console", "", 500, 700)
If X = "timer1" Then
If Timer1.Enabled = True Then
Timer1.Enabled = False
MsgBox "timer1 OFF", , "console prompt"
Else
Timer1.Enabled = True
MsgBox " timer1 ON", , "console prompt"
End If
End If
If X = "timer2" Then
If Timer2.Enabled = True Then
Timer2.Enabled = False
MsgBox "timer2 OFF", , "console prompt"
Else
Timer2.Enabled = True
MsgBox "timer2 ON", , "console prompt"
End If
End If
If X = "timer3" Then
If Timer3.Enabled = True Then
Timer3.Enabled = False
MsgBox "timer3 OFF", , "console prompt"
Else
Timer3.Enabled = True
MsgBox "timer3 ON", , "console prompt"
End If
End If
If X = "timer4" Then
If Timer4.Enabled = True Then
Timer4.Enabled = False
MsgBox "timer4 OFF", , "console prompt"
Else
Timer4.Enabled = True
MsgBox "timer4 ON", , "console prompt"
End If
End If
If X = "timer5" Then
If Timer5.Enabled = True Then
Timer5.Enabled = False
MsgBox "timer5 OFF", , "console prompt"
Else
Timer5.Enabled = True
MsgBox "timer5 ON", , "console prompt"
End If
End If
If X = "timer6" Then
If Timer6.Enabled = True Then
Timer6.Enabled = False
MsgBox "timer6 OFF", , "console prompt"
Else
Timer6.Enabled = True
MsgBox "timer6 ON", , "console prompt"
End If
End If
If X = "gk" Then
Y = InputBox("Level", "goalkeeper difficulty(def = 60) ", , 500, 700)
If Y = Empty Then
Exit Sub
End If
If Y > 100 Then
MsgBox "bara kamal karna laga hai!bachi bhi goal karle!", , "console prompt"
End If
Label7.Caption = Y
End If
If X = "gk2" Then
z = InputBox("Level2", "goalkeeper jump difficulty(def = 30) ", , 500, 700)
If z = Empty Then
Exit Sub
End If
If z > 100 Then
MsgBox "bara kamal karna laga hai!bachi bhi goal karle!", , "console prompt"
End If
Label8.Caption = z
End If
If X = "nogk" And Shape3.Visible = True Then
Timer7.Enabled = False
Shape3.Visible = False
Timer10.Enabled = False
Shape3.Top = 100000
MsgBox "NoGoalkeeper ON", , "console prompt"
Label11.Caption = "0"
Exit Sub
End If
If X = "nogk" And Shape3.Visible = False Then
Timer7.Enabled = True
Timer10.Enabled = False
Shape3.Visible = True
MsgBox "NoGoalkeeper OFF", , "console prompt"
Label11.Caption = "1"
End If
If X = "change_song" Then
MsgBox "Track Change succesfull", , "console prompt"


If WindowsMediaPlayer1.URL = "C:\RECORDER\REC025.MP3" Then
WindowsMediaPlayer1.URL = "C:\RECORDER\REC024.MP3"
Exit Sub
End If
If WindowsMediaPlayer1.URL = "C:\RECORDER\REC024.MP3" Then
WindowsMediaPlayer1.URL = "C:\RECORDER\favela rock afrobots.MP3"
Exit Sub
End If
If WindowsMediaPlayer1.URL = "C:\RECORDER\favela rock afrobots.mp3" Then
WindowsMediaPlayer1.URL = "C:\RECORDER\happy as can be.MP3"
Exit Sub
End If
If WindowsMediaPlayer1.URL = "C:\RECORDER\happy as can be.MP3" Then
WindowsMediaPlayer1.URL = "C:\RECORDER\REC029.MP3"
Exit Sub
End If
If WindowsMediaPlayer1.URL = "C:\RECORDER\REC029.mp3" Then
WindowsMediaPlayer1.URL = "C:\RECORDER\REC025.MP3"
Exit Sub
End If
End If
If X = "wind_speed" Then
a = InputBox("Wind Speed", "Environment", , 500, 700)
Text4.Text = Val(a)
MsgBox "" & Text4.Text & " is now the wind speed""", , "console prompt"
End If
If X = "wind_direction" Then
b = InputBox("Wind Direction", "Environment", , 500, 700)
If b = 1 Then
Text5.Text = "right"
MsgBox "" & Text5.Text & " is now the wind speed""", , "console prompt"
End If
If b = 0 Then
Text5.Text = "left"
MsgBox "" & Text5.Text & " is now the wind speed""", , "console prompt"
End If
End If
If X = "gravity" Then
C = InputBox("Gravity (def = 1)", "Environment", , 500, 700)
Label14.Caption = Val(C)
MsgBox "" & Val(Label14.Caption) & " is now the gravity""", , "console prompt"
End If
If X = "friction" Then
d = InputBox("Friction (def = 0)", "Environment", , 500, 700)
Label16.Caption = Val(d)
MsgBox "" & Val(Label16.Caption) & " is now the friction""", , "console prompt"
End If
End Sub



Private Sub Form_Load()
For X = 1 To 10
Next X
X = Rnd * 50
Text4.Text = X
End Sub

Private Sub Command4_Click()
MsgBox "As the power increases then the spin decreases and the shot becomes more fast", , "Balancing Power And Spin"
MsgBox "Hint: The keeper is right handed so he has a weak left hand grip and is a little more likely to miss saves on left side", , "Aiming Tips"
End Sub

Private Sub Command5_Click()
Unload Form1
Load Form1
Form1.Show
End Sub

Private Sub Command6_Click()
If WindowsMediaPlayer1.URL = "C:\RECORDER\REC029.mp3" Then
WindowsMediaPlayer1.URL = "C:\RECORDER\REC025.MP3"
Exit Sub
End If
End Sub

Private Sub Command7_Click()
Timer10.Enabled = False
End Sub

Private Sub Command8_Click()
Timer10.Enabled = False
End Sub

Private Sub windowsmediaplayer1_EndOfStream(ByVal Result As Long)
If WindowsMediaPlayer1.URL = "C:\RECORDER\REC029.MP3" Then
WindowsMediaPlayer1.URL = "C:\RECORDER\REC025.MP3"
Exit Sub
End If
If WindowsMediaPlayer1.URL = "C:\RECORDER\REC025.MP3" Then
WindowsMediaPlayer1.URL = "C:\RECORDER\REC024.MP3"
Exit Sub
End If
If WindowsMediaPlayer1.URL = "C:\RECORDER\REC024.MP3" Then
WindowsMediaPlayer1.URL = "C:\RECORDER\REC029.MP3"
Exit Sub
End If
End Sub

Private Sub Text2_Change()
If Val(Text2.Text) = 0 Then
Text2.Text = 0
End If
End Sub

Private Sub Text3_Change()
If Val(Text3.Text) = 0 Then
Text3.Text = 0
End If
End Sub

Private Sub Text4_Change()
If Val(Text4.Text) > 0 Then
Text5.Text = "Right"
Else
Text5.Text = "Left"
End If
End Sub

Private Sub Timer1_Timer()
X = Text1.Text
If X > 100 Then
X = 100
End If
If Shape1.Height < 375 Then
Timer1.Enabled = False
Timer3.Enabled = True
Exit Sub
End If
Shape1.Height = Val(Shape1.Height) - X - 20
If Val(Text2.Text) > 0 Then
Shape1.Left = Val(Shape1.Left) + Val(Text2.Text) + Val(Text4.Text)
End If
If Val(Text2.Text) < 0 Then
Shape1.Left = Val(Shape1.Left) + Val(Text2.Text) + Val(Text4.Text)
End If
If Text3.Text > 0 Then
Shape1.Top = Val(Shape1.Top) - Val(Text3.Text)
Else
Shape1.Top = Val(Shape1.Top) + Val(Text3.Text)
End If
End Sub

Private Sub Timer10_Timer()
X = 4500 + Val(Text2.Text) * 5 - Val(Shape3.Left)
X = X / 100
If Val(Text2.Text) > 0 Then
Shape3.Left = Val(Shape3.Left) + X + 10
Else
Shape3.Left = Val(Shape3.Left) + X - 10
End If
Y = 2640 - Val(Text3.Text) - Val(Shape3.Top)
Y = Y / 80
Shape3.Top = Val(Shape3.Top) + Y
If Shape3.Top < 1500 Then
Shape3.Top = 2400
Exit Sub
End If
End Sub

Private Sub Timer11_Timer()
Label9.Top = Shape3.Top
Label9.Left = Shape3.Left
End Sub

Private Sub Timer12_Timer()
Label10.Caption = Val(Shape1.Height) - 350
Label10.Top = Val(Shape1.Top) + 200
Label10.Left = Val(Shape1.Left) + 100
End Sub

Private Sub Timer13_Timer()
If Shape5.Top >= 9720 Then
Shape5.Top = "9120"
Timer14.Enabled = True
Timer13.Enabled = False
Else
Shape5.Top = Shape1.Top + 6120
End If
End Sub

Private Sub Timer14_Timer()
Shape5.Top = Shape1.Top + 5500
End Sub

Private Sub Timer15_Timer()
X = 855 - Shape1.Height
X = X * 10
Shape5.Left = Label5.Left + X
End Sub

Private Sub Timer16_Timer()
Shape6.Top = Shape3.Top + 5760
If Shape6.Top >= 8260 Then
Timer16.Enabled = False
End If
End Sub

Private Sub Timer17_Timer()
X = Shape3.Top + 5760
If X <= 8160 Then
Timer16.Enabled = True
End If
End Sub

Private Sub Timer18_Timer()
If Shape5.Top >= 9240 Then
Timer1.Interval = Timer1.Interval + Val(Label16.Caption)
End If
If Timer1.Interval > 500 Then
Timer1.Enabled = False
MsgBox "ball stopped"
Timer18.Enabled = False
End If
End Sub

Private Sub Timer2_Timer()
If Shape1.Top > 3600 Then
Exit Sub
End If
If Shape1.Height < 375 Then
Label4.Caption = 1
Timer2.Enabled = False
End If
Label4.Caption = Val(Label4.Caption) + Val(Label14.Caption)
Shape1.Top = Val(Shape1.Top) + Val(Label4.Caption)
End Sub

Private Sub Timer3_Timer()
If Shape1.Height < 375 And Shape1.Top > 1200 And Shape1.Left < 6840 And Shape1.Left > 2000 Then
MsgBox "GOAL"
Timer3.Enabled = False
Timer7.Enabled = False
Shape1.Height = Val(Shape1.Height) - 20
End If
End Sub

Private Sub Timer5_Timer()
X = Val(Label5.Caption) / 2
If Label5.Caption < Shape1.Top And Shape1.Top > 3400 Then
Label6.Caption = Val(Label6.Caption) + 200
Shape1.Top = Shape1.Top - X + Val(Label6.Caption)
If Shape1.Top > 3600 Then
Shape1.Top = 3600
Timer5.Enabled = False
End If
Label5.Caption = 10000
End If
End Sub

Private Sub Timer6_Timer()
If Label5.Caption > Shape1.Top Then
Label5.Caption = Shape1.Top
End If
End Sub

Private Sub Timer7_Timer()
X = Val(Shape1.Left) - Val(Shape3.Left)
X = X / Val(Label7.Caption)
Shape3.Left = Val(Shape3.Left) + X
Y = Val(Shape1.Top) - Val(Shape3.Top)
Y = Y / Val(Label8.Caption)
Shape3.Top = Val(Shape3.Top) + Y
If Shape3.Top < 1500 Then
Shape3.Top = 2400
Exit Sub
End If
End Sub

Private Sub Timer8_Timer()
X = Val(Shape3.Left) - Val(Shape1.Left)
Y = Val(Shape3.Top) - Val(Shape1.Top)
If Shape1.Height < 400 And X < 350 And X > -300 And Y < 300 Then
Shape1.Height = 375
MsgBox "goalkeeper saves!"
Timer8.Enabled = False
Timer1.Enabled = False
Timer2.Enabled = False
End If
End Sub

Private Sub Timer9_Timer()
Text1.Text = Val(Text1.Text) * 1.01 + 1
If Text1.Text > 100 Then
Text1.Text = 1
End If
End Sub

Private Sub Windowswindowsmediaplayer1_OpenStateChange(ByVal NewState As Long)

End Sub
