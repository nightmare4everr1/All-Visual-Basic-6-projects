VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "0"
   ClientHeight    =   9690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14070
   LinkTopic       =   "Form1"
   ScaleHeight     =   9690
   ScaleWidth      =   14070
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   4935
      Left            =   4800
      TabIndex        =   10
      Top             =   4080
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.Timer Timer7 
      Left            =   3720
      Top             =   1320
   End
   Begin VB.CommandButton Command7 
      Caption         =   "stop car2"
      Height          =   375
      Left            =   720
      TabIndex        =   9
      Top             =   7200
      Width           =   1455
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2280
      Top             =   6720
   End
   Begin VB.CommandButton Command6 
      Caption         =   "move car2"
      Height          =   375
      Left            =   600
      TabIndex        =   8
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Timer Timer5 
      Interval        =   1
      Left            =   1440
      Top             =   3240
   End
   Begin VB.Timer Timer4 
      Interval        =   1
      Left            =   600
      Top             =   3240
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2040
      Top             =   1800
   End
   Begin VB.CommandButton Command5 
      Caption         =   "3rd gear"
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "2nd gear"
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   2520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "1st gear"
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   2520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   3120
      Top             =   840
   End
   Begin VB.CommandButton Command2 
      Caption         =   "brake"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2520
      Top             =   840
   End
   Begin VB.CommandButton Command1 
      Caption         =   "accelerate"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "1.3"
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "0.04"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "140"
      BeginProperty Font 
         Name            =   "HL2cross"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   4200
      TabIndex        =   1
      Top             =   2520
      Width           =   5295
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000003&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   4560
      Shape           =   4  'Rounded Rectangle
      Top             =   3840
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   720
      Shape           =   4  'Rounded Rectangle
      Top             =   3840
      Width           =   135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer1.Enabled = True
Timer1.Tag = 1
End Sub

Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer1.Enabled = False
Timer1.Tag = 0
End Sub

Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer3.Enabled = True
End Sub

Private Sub Command2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer3.Enabled = False
End Sub

Private Sub Command3_Click()
Label2.Caption = "0.04"
Label3.Caption = "1.3"
End Sub

Private Sub Command4_Click()
Label2.Caption = "0.03"
Label3.Caption = "1.5"
End Sub

Private Sub Command5_Click()
Label2.Caption = "0.02"
Label3.Caption = "1.7"
End Sub

Private Sub Command6_Click()
Timer6.Enabled = True
End Sub

Private Sub Command7_Click()
Timer6.Enabled = False
End Sub

Private Sub Timer1_Timer()
If Label1.Caption = 0 Then
Label1.Caption = 1
End If
X = Label1.Caption
X = Label3.Caption / X
Label1.Caption = Label1.Caption + X
End Sub

Private Sub Timer2_Timer()
If Label1.Caption <= 0 Then
Label1.Caption = 0
Exit Sub
End If
Label1.Caption = Label1.Caption - Label2.Caption
End Sub

Private Sub Timer3_Timer()
If Label1.Caption <= 0 Then
Label1.Caption = "0"
Exit Sub
End If
If Label1.Caption <= 10 Then
Label1.Caption = "0"
End If
Label1.Caption = Label1.Caption - 0.2
End Sub

Private Sub Timer4_Timer()
Shape1.Left = Shape1.Left + Label1.Caption / 3.6
End Sub

Private Sub Timer5_Timer()
If Label1.Caption <= 0 Then
Exit Sub
End If
X = Shape2.Left - Shape1.Left - 300
X = X / 100
s = Label1.Caption / 3.6
t = X / s
List1.AddItem (t & "     " & X)
If t < 1 Then
Label1.Caption = Label1.Caption - 1
List1.AddItem (aaaaaaaaaaaaaa)
Timer1.Enabled = False
Else
If Timer1.Tag = "1" Then
Timer1.Enabled = True
End If
End If
End Sub

Private Sub Timer6_Timer()
Shape2.Left = Shape2.Left + 5
End Sub

Private Sub Timer7_Timer()
Label1.Caption = Int(Label1.Caption)
End Sub
