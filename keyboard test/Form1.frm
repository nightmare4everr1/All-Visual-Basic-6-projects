VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9630
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13965
   LinkTopic       =   "Form1"
   ScaleHeight     =   9630
   ScaleWidth      =   13965
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5160
      Top             =   2880
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1320
      Top             =   240
   End
   Begin VB.TextBox Text1 
      Height          =   1095
      Left            =   3600
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1440
      Width           =   4935
   End
   Begin WMPLibCtl.WindowsMediaPlayer Mediaplayer1 
      Height          =   735
      Left            =   -1080
      TabIndex        =   3
      Top             =   2520
      Width           =   1815
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
      _cx             =   3201
      _cy             =   1296
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   135
      Left            =   480
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Enter text below"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3720
      TabIndex        =   0
      Top             =   360
      Width           =   5535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If Timer2.Enabled = True Then
Timer2.Tag = "1"
Exit Sub
End If
If Timer2.Enabled = False Then
Timer2.Enabled = True
End If
Select Case Label2.Caption
Case Is >= 9
Mediaplayer1.URL = "C:\Microsoft Visual Studio\VB98\keyboard test\keyboard1_clicks.wav"
Case Is >= 7.5
Mediaplayer1.URL = "C:\Microsoft Visual Studio\VB98\keyboard test\keyboard2_clicks.wav"
Case Is >= 6
Mediaplayer1.URL = "C:\Microsoft Visual Studio\VB98\keyboard test\keyboard3_clicks.wav"
Case Is >= 4.5
Mediaplayer1.URL = "C:\Microsoft Visual Studio\VB98\keyboard test\keyboard4_clicks.wav"
Case Is >= 3
Mediaplayer1.URL = "C:\Microsoft Visual Studio\VB98\keyboard test\keyboard5_clicks.wav"
Case Is >= 1.5
Mediaplayer1.URL = "C:\Microsoft Visual Studio\VB98\keyboard test\keyboard6_clicks.wav"
Case Is >= 0
Mediaplayer1.URL = "C:\Microsoft Visual Studio\VB98\keyboard test\keyboard7_clicks_enter.wav"
End Select
End Sub

Private Sub Timer1_Timer()
Label2.Caption = Rnd * 10
End Sub

Private Sub Timer2_Timer()
If Timer2.Tag <> "1" Then
Timer2.Enabled = False
Exit Sub
End If
'If MediaPlayer1.playState <> wmppsPlaying Then
'Select Case Label2.Caption
'Case Is >= 6.6
'MediaPlayer1.URL = "C:\Microsoft Visual Studio\VB98\keyboard test\keyboard_fast1_1second.wav"
'Case Is >= 3.3
'MediaPlayer1.URL = "C:\Microsoft Visual Studio\VB98\keyboard test\keyboard_fast2_1second.wav"
'Case Is >= 0
'MediaPlayer1.URL = "C:\Microsoft Visual Studio\VB98\keyboard test\keyboard_fast3_1second.wav"
'End Select
'End If
Timer2.Enabled = False
Timer2.Tag = "0"
End Sub
