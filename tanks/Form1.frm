VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9705
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14070
   LinkTopic       =   "Form1"
   ScaleHeight     =   9705
   ScaleWidth      =   14070
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer5 
      Interval        =   1
      Left            =   9000
      Top             =   6960
   End
   Begin VB.Timer Timer4 
      Interval        =   1
      Left            =   2160
      Top             =   480
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Text            =   "1"
      Top             =   4440
      Width           =   975
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3720
      Top             =   2880
   End
   Begin VB.Timer Timer2 
      Left            =   5520
      Top             =   5760
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1320
      Top             =   6840
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "turn"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   21
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "player"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   20
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label18 
      Caption         =   "player2 score"
      Height          =   375
      Left            =   5400
      TabIndex        =   19
      Top             =   8640
      Width           =   1695
   End
   Begin VB.Label Label17 
      Caption         =   "player1 score"
      Height          =   375
      Left            =   3240
      TabIndex        =   18
      Top             =   8640
      Width           =   1215
   End
   Begin VB.Label Label16 
      Caption         =   "0"
      Height          =   495
      Left            =   5400
      TabIndex        =   17
      Top             =   9000
      Width           =   1695
   End
   Begin VB.Label Label15 
      Caption         =   "0"
      Height          =   375
      Left            =   3240
      TabIndex        =   16
      Top             =   9120
      Width           =   1215
   End
   Begin VB.Label Label14 
      Caption         =   "Label14"
      Height          =   615
      Left            =   7200
      TabIndex        =   15
      Top             =   6720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label13 
      Caption         =   "2"
      Height          =   495
      Left            =   5760
      TabIndex        =   14
      Top             =   840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label11 
      Caption         =   "1"
      Height          =   375
      Left            =   4440
      TabIndex        =   13
      Top             =   840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Shape3 
      Height          =   495
      Left            =   10200
      Top             =   8040
      Width           =   855
   End
   Begin VB.Line Line4 
      X1              =   11280
      X2              =   10680
      Y1              =   7680
      Y2              =   8160
   End
   Begin VB.Shape Shape4 
      Height          =   135
      Left            =   11280
      Shape           =   3  'Circle
      Top             =   5400
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "2"
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
      Left            =   5160
      TabIndex        =   12
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Power of the rocket thrusters"
      Height          =   495
      Left            =   360
      TabIndex        =   11
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   480
      Visible         =   0   'False
      Width           =   1695
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   615
      Left            =   7800
      TabIndex        =   9
      Top             =   1320
      Visible         =   0   'False
      Width           =   2295
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
      _cx             =   4048
      _cy             =   1085
   End
   Begin VB.Label Label8 
      Caption         =   "reset"
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   495
      Left            =   4200
      TabIndex        =   6
      Top             =   3360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   495
      Left            =   2880
      TabIndex        =   5
      Top             =   3360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Label5"
      Height          =   375
      Left            =   5280
      TabIndex        =   4
      Top             =   6720
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   14160
      Y1              =   8520
      Y2              =   8520
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   7920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   5760
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   5760
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Shape Shape2 
      Height          =   255
      Left            =   3120
      Shape           =   3  'Circle
      Top             =   7560
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fire"
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   3120
      X2              =   2520
      Y1              =   7680
      Y2              =   8160
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Left            =   2040
      Top             =   8040
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Circle (100, 100), Radius
End Sub

Private Sub Form_Click()
'Dim CX, CY, Radius, Limit   ' Declare variable.
' ScaleMode = 3   ' Set scale to pixels.
    'CX = ScaleWidth / 2 ' Set X position.
    'CY = ScaleHeight / 2    ' Set Y position.
    'If CX > CY Then Limit = CY Else Limit = CX
    'For Radius = 0 To Limit ' Set radius.
      '  Circle (CX, CY), 500, RGB(Rnd * 255, Rnd * 255, Rnd * 255)
    'Next Radius
    'Circle (1000, 100), Radius
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyUp:
If Label12.Caption = "1" Then

If Line1.Y1 >= Line1.Y2 And Line1.X1 < Line1.X2 Then
Exit Sub
End If
Line1.X1 = Line1.X1 - 10
If Line1.X1 < Line1.X2 Then
Line1.Y1 = Line1.Y1 + 10
Else
Line1.Y1 = Line1.Y1 - 10
End If

End If

If Label12.Caption = "2" Then

If Line4.Y1 >= Line4.Y2 And Line4.X1 < Line4.X2 Then
Exit Sub
End If
Line4.X1 = Line4.X1 - 10
If Line4.X1 < Line4.X2 Then
Line4.Y1 = Line4.Y1 + 10
Else
Line4.Y1 = Line4.Y1 - 10
End If
End If

Case vbKeyDown
If Label12.Caption = "1" Then

If Line1.Y1 >= Line1.Y2 And Line1.X1 > Line1.X2 Then
Exit Sub
End If
Line1.X1 = Line1.X1 + 10
If Line1.X1 > Line1.X2 Then
Line1.Y1 = Line1.Y1 + 10
Else
Line1.Y1 = Line1.Y1 - 10
End If
Else

If Line4.Y1 >= Line4.Y2 And Line4.X1 > Line4.X2 Then
Exit Sub
End If
Line4.X1 = Line4.X1 + 10
If Line4.X1 > Line4.X2 Then
Line4.Y1 = Line4.Y1 + 10
Else
Line4.Y1 = Line4.Y1 - 10
End If
End If

Case vbKeyRight
If Label12.Caption = "1" Then
x = Shape1.Left - Shape3.Left
If x < 0 Then
x = x * -1
End If
If x <= 2000 Then
Exit Sub
End If
Shape1.Left = Shape1.Left + 500
Line1.X1 = Line1.X1 + 500
Line1.X2 = Line1.X2 + 500
Else

Shape3.Left = Shape3.Left + 500
Line4.X1 = Line4.X1 + 500
Line4.X2 = Line4.X2 + 500


End If

Case vbKeyLeft
If Label12.Caption = "1" Then

Shape1.Left = Shape1.Left - 500
Line1.X1 = Line1.X1 - 500
Line1.X2 = Line1.X2 - 500
Else
x = Shape3.Left - Shape1.Left
If x < 0 Then
x = x * -1
End If
If x <= 2000 Then
Exit Sub
End If
Shape3.Left = Shape3.Left - 500
Line4.X1 = Line4.X1 - 500
Line4.X2 = Line4.X2 - 500
End If
End Select

End Sub

Private Sub Form_Load()
KeyPreview = True
End Sub

Private Sub Label1_Click()
Select Case Label9.Caption
Case Is >= 5
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\mortar_fire1.wav"
Case Is >= 0
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\grenade_launcher1.wav"
End Select
WindowsMediaPlayer1.Controls.play
If Label12.Caption = "2" Then
Timer3.Enabled = True
Label14.Caption = "fire"
Label7.Caption = Val(Text1.Text) / 10
Shape4.Visible = True
Label12.Caption = "1"
Exit Sub
End If
If Label12.Caption = "1" Then
Label12.Caption = "2"
Label4.Caption = "fire"
Timer3.Enabled = True
Label7.Caption = Val(Text1.Text) / 10
Shape2.Visible = True
End If
Select Case Label9.Caption
Case Is >= 5
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\mortar_fire1.wav"
Case Is >= 0
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\grenade_launcher1.wav"
End Select
WindowsMediaPlayer1.Controls.play

End Sub

Private Sub Label11_Click()
Label12.Caption = "1"
End Sub

Private Sub Label13_Click()
Label12.Caption = "2"
End Sub

Private Sub Label8_Click()
Unload Form1
Load Form1
Form1.Show
End Sub

Private Sub Text1_Change()
If Val(Text1.Text) = Empty Or Val(Text1.Text) < 1 Then
Text1.Text = "1"
End If
If Val(Text1.Text) > 30 Then
Text1.Text = 30
End If

End Sub

Private Sub Timer1_Timer()

If Shape2.Left > 15000 Or Shape2.Left < 0 Then
Timer3.Enabled = False
Shape2.Visible = False
Label4.Caption = "stop"
Label6.Caption = "1"
End If
If Label4.Caption <> "fire" Then
Shape2.Top = Line1.Y1
Shape2.Left = Line1.X1
If Line1.X1 > Line1.X2 Then
Label2.Caption = Line1.Y1 - 7080
Label2.Caption = Label2.Caption / 10
Else
Label2.Caption = Line1.Y1 - 7080
Label2.Caption = Label2.Caption / 10
Label2.Caption = Label2.Caption * -1
End If
End If
If Label4.Caption = "fire" Then
Shape2.Left = Shape2.Left + Label2.Caption
x = Label2.Caption - 110
Shape2.Top = Shape2.Top + x
End If


If Shape2.Top >= Shape3.Top And Shape2.Left >= Shape3.Left And Shape2.Left < Shape3.Left + 1000 Or Shape2.Top + Shape2.Height >= 8400 And Shape2.Left >= Shape3.Left - Shape2.Width And Shape2.Left <= Shape3.Left + Shape2.Width Then
'MsgBox "you win"
Label15.Caption = Label15.Caption + 1
Select Case Label9.Caption
Case Is >= 9
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode_1.wav"
Case Is >= 8
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode_2.wav"
Case Is >= 6.6
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode_3.wav"
Case Is >= 5.5
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode_4.wav"
Case Is >= 3
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode_5.wav"
Case Is >= 2
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode_7.wav"
Case Is >= 1
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode_8.wav"
Case Is >= 0
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode_9.wav"
End Select
Timer3.Enabled = False
Shape2.Visible = False
Label4.Caption = "stop"
Label6.Caption = "1"
End If
If Shape2.Top + Shape2.Height > 9000 Then
'MsgBox "boom"
Timer3.Enabled = False
Shape2.Visible = False
Label4.Caption = "stop"
Label6.Caption = "1"
Select Case Label9.Caption
Case Is >= 9
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode3.wav"
Case Is >= 8
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode4.wav"
Case Is >= 6.6
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode5.wav"
Case Is >= 5.5
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode3.wav"
Case Is >= 3
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode4.wav"
Case Is >= 2
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode5.wav"
Case Is >= 1
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode3.wav"
Case Is >= 0
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode4.wav"
End Select
End If


End Sub

Private Sub Timer3_Timer()
Label6.Caption = Val(Label6.Caption) + Val(Label7.Caption)
Shape2.Top = Shape2.Top + Val(Label6.Caption)
Shape4.Top = Shape4.Top + Val(Label6.Caption)
End Sub

Private Sub Timer4_Timer()
Label9.Caption = Rnd * 10
End Sub

Private Sub Timer5_Timer()

If Shape4.Left > 15000 Or Shape4.Left < 0 Then
Timer3.Enabled = False
Shape4.Visible = False
Label14.Caption = "stop"
Label6.Caption = "1"
End If
If Label14.Caption <> "fire" Then
Shape4.Top = Line4.Y1
Shape4.Left = Line4.X1
If Line4.X1 > Line4.X2 Then
Label3.Caption = Line4.Y1 - 7080
Label3.Caption = Label3.Caption / 10
Else
Label3.Caption = Line4.Y1 - 7080
Label3.Caption = Label3.Caption / 10
Label3.Caption = Label3.Caption * -1
End If
End If
If Label14.Caption = "fire" Then
Shape4.Left = Shape4.Left + Label3.Caption
If Line4.X1 > Line4.X2 Then
x = Label3.Caption - 110
Shape4.Top = Shape4.Top + x
Else
x = Label3.Caption + 110
Shape4.Top = Shape4.Top - x
End If

End If
If Shape4.Top >= Shape1.Top And Shape4.Left <= Shape1.Left + 1000 And Shape4.Left > Shape1.Left Or Shape4.Top + Shape4.Height >= 8400 And Shape4.Left <= Shape1.Left - Shape4.Width And Shape4.Left >= Shape1.Left + Shape4.Width Then
'MsgBox "you win"
Label16.Caption = Label16.Caption + 1
Select Case Label9.Caption
Case Is >= 9
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode_1.wav"
Case Is >= 8
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode_2.wav"
Case Is >= 6.6
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode_3.wav"
Case Is >= 5.5
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode_4.wav"
Case Is >= 3
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode_5.wav"
Case Is >= 2
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode_7.wav"
Case Is >= 1
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode_8.wav"
Case Is >= 0
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode_9.wav"
End Select
Timer3.Enabled = False
Shape4.Visible = False
Label14.Caption = "stop"
Label6.Caption = "1"
End If
If Shape4.Top > 8400 Then
'MsgBox "boom"
Timer3.Enabled = False
Shape4.Visible = False
Label14.Caption = "stop"
Label6.Caption = "1"
Select Case Label9.Caption
Case Is >= 9
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode3.wav"
Case Is >= 8
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode4.wav"
Case Is >= 6.6
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode5.wav"
Case Is >= 5.5
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode3.wav"
Case Is >= 3
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode4.wav"
Case Is >= 2
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode5.wav"
Case Is >= 1
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode3.wav"
Case Is >= 0
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode4.wav"
End Select
End If
End Sub

Private Sub Timer6_Timer()
If Shape2.Top >= Shape3.Top And Shape2.Left >= Shape3.Left Then
'MsgBox "you win"
Select Case Label9.Caption
Case Is >= 9
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode_1.wav"
Case Is >= 8
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode_2.wav"
Case Is >= 6.6
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode_3.wav"
Case Is >= 5.5
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode_4.wav"
Case Is >= 3
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode_5.wav"
Case Is >= 2
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode_7.wav"
Case Is >= 1
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode_8.wav"
Case Is >= 0
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode_9.wav"
End Select
Timer3.Enabled = False
Shape2.Visible = False
Label4.Caption = "stop"
Label6.Caption = "1"
End If
If Shape2.Top > 8400 Then
'MsgBox "boom"
Timer3.Enabled = False
Shape2.Visible = False
Label4.Caption = "stop"
Label6.Caption = "1"
Select Case Label9.Caption
Case Is >= 9
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode3.wav"
Case Is >= 8
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode4.wav"
Case Is >= 6.6
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode5.wav"
Case Is >= 5.5
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode3.wav"
Case Is >= 3
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode4.wav"
Case Is >= 2
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode5.wav"
Case Is >= 1
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode3.wav"
Case Is >= 0
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode4.wav"
End Select
End If






If Shape4.Top >= Shape1.Top And Shape4.Left <= Shape1.Left + 1000 Then
'MsgBox "you win"
Select Case Label9.Caption
Case Is >= 9
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode_1.wav"
Case Is >= 8
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode_2.wav"
Case Is >= 6.6
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode_3.wav"
Case Is >= 5.5
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode_4.wav"
Case Is >= 3
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode_5.wav"
Case Is >= 2
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode_7.wav"
Case Is >= 1
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode_8.wav"
Case Is >= 0
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode_9.wav"
End Select
Timer3.Enabled = False
Shape4.Visible = False
Label4.Caption = "stop"
Label6.Caption = "1"
End If
If Shape4.Top > 8400 Then
'MsgBox "boom"
Timer3.Enabled = False
Shape4.Visible = False
Label4.Caption = "stop"
Label6.Caption = "1"
Select Case Label9.Caption
Case Is >= 9
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode3.wav"
Case Is >= 8
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode4.wav"
Case Is >= 6.6
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode5.wav"
Case Is >= 5.5
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode3.wav"
Case Is >= 3
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode4.wav"
Case Is >= 2
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode5.wav"
Case Is >= 1
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode3.wav"
Case Is >= 0
WindowsMediaPlayer1.URL = "C:\Visual BASIC 5.0 (Ent. Edition)\VB\tanks\explode4.wav"
End Select
End If
End Sub
