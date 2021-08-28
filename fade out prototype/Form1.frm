VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9600
   ClientLeft      =   -435
   ClientTop       =   915
   ClientWidth     =   13935
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9600
   ScaleWidth      =   13935
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   11040
      TabIndex        =   15
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "change password"
      Height          =   495
      Left            =   9960
      TabIndex        =   14
      Top             =   6480
      Width           =   975
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   3360
      Top             =   5520
   End
   Begin VB.CommandButton Command5 
      Caption         =   "enter password"
      Height          =   615
      Left            =   4200
      TabIndex        =   11
      Top             =   6600
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   6000
      TabIndex        =   10
      Tag             =   "178296"
      Text            =   "0"
      Top             =   6480
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      Caption         =   "hack tool"
      Height          =   975
      Left            =   4080
      TabIndex        =   9
      Top             =   5400
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   600
      TabIndex        =   8
      Top             =   3720
      Width           =   975
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   8640
      Top             =   8400
   End
   Begin VB.CommandButton Command4 
      Caption         =   "focus in"
      Height          =   735
      Left            =   6360
      TabIndex        =   6
      Top             =   7320
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "focus out"
      Height          =   975
      Left            =   6480
      TabIndex        =   5
      Top             =   8280
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      ClipControls    =   0   'False
      Height          =   2175
      Left            =   2400
      TabIndex        =   4
      Top             =   7200
      Width           =   3855
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Main Menu"
         Height          =   735
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   8640
      Top             =   7440
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   1200
      Top             =   2880
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000C&
      Caption         =   "Processing.."
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   1800
      TabIndex        =   13
      Top             =   2160
      Visible         =   0   'False
      Width           =   7575
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   735
      Left            =   7200
      TabIndex        =   12
      Top             =   5520
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   3000
      Left            =   2520
      Picture         =   "Form1.frx":0000
      Top             =   960
      Width           =   7200
   End
   Begin VB.Label Label4 
      Caption         =   "150"
      Height          =   855
      Left            =   480
      TabIndex        =   3
      Top             =   7440
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "150"
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "150"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   6600
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim RED, I, RGBValue, Form1
RED = RGB(255, 0, 0)    ' Return the value for Red.
I = 75  ' Initialize offset.
RGBValue = RGB(I, 64 + I, 128 + I)  ' Same as RGB(75, 139, 203).
Form1.BackColor = RGB(255, 0, 0) ' Set the Color property of
' MyObject to Red.
End Sub

Private Sub Command2_Click()
Label6.Caption = "0"
Label7.Visible = True
Timer4.Enabled = True
End Sub

Private Sub Command3_Click()
Timer3.Enabled = True
Timer2.Enabled = False
End Sub

Private Sub Command4_Click()
Timer2.Enabled = True
Timer3.Enabled = False
End Sub

Private Sub Command5_Click()

If Text1.Text = Text1.Tag Then
MsgBox "welcome"
End If
End Sub

Private Sub Command6_Click()
Text1.Tag = Text2.Text
Text2.Text = Empty
End Sub

Private Sub Form_Load()
Frame1.BackColor = RGB(150, 150, 150)
End Sub

Private Sub Timer1_Timer()
Image1.Left = Image1.Left + 10
Image1.Top = Image1.Top + 10
Label1.Caption = Val(Label1.Caption) + 1
If Label1.Caption >= 50 Then
Timer1.Enabled = False
End If

End Sub

Private Sub Timer2_Timer()
If Label2.Caption < 150 Then
Label2.Caption = Val(Label2.Caption) + 2
Label3.Caption = Val(Label3.Caption) + 2
Label4.Caption = Val(Label4.Caption) + 2
a = Label2.Caption
b = Label3.Caption
c = Label4.Caption
Frame1.BackColor = RGB(a, b, c)
Else
Timer2.Enabled = False
End If


End Sub

Private Sub Timer3_Timer()
If Label2.Caption >= 100 Then
Label2.Caption = Val(Label2.Caption) - 2
Label3.Caption = Val(Label3.Caption) - 2
Label4.Caption = Val(Label4.Caption) - 2
a = Label2.Caption
b = Label3.Caption
c = Label4.Caption
Frame1.BackColor = RGB(a, b, c)
Else
Timer3.Enabled = False
End If
End Sub

Private Sub Timer4_Timer()
Do While Val(Label6.Caption) <> Val(Text1.Tag)
Label6.Caption = Val(Label6.Caption) + 1
If Label6.Caption = Val(Text1.Tag) Then
Label7.Visible = False
Timer4.Enabled = False
End If
Loop
End Sub
