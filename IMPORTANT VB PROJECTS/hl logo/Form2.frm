VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H80000008&
   Caption         =   "Form2"
   ClientHeight    =   2415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5445
   LinkTopic       =   "Form2"
   ScaleHeight     =   2415
   ScaleWidth      =   5445
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   960
      Top             =   240
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0080FFFF&
      Height          =   615
      Left            =   2400
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H8000000D&
      Caption         =   "play a test"
      Height          =   615
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000D&
      Caption         =   "create a test /see old papers"
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   2160
      TabIndex        =   5
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Caption         =   $"Form2.frx":0000
      ForeColor       =   &H8000000C&
      Height          =   1215
      Left            =   -5460
      TabIndex        =   4
      Top             =   0
      Width           =   5415
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   255
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = Text1.Tag Then
form3.Command1.Visible = True
form3.Command2.Visible = True
form3.Command3.Visible = True
form3.Command4.Visible = True
form3.Command5.Visible = True
form3.Command6.Visible = True
form3.Command7.Visible = True
form3.Command8.Visible = True
form3.Command10.Visible = True
form3.Check1.Visible = True
form3.Check2.Visible = True
form3.Check3.Visible = True
form3.Check4.Visible = True
form3.Text1.Visible = True
form3.Label2.Visible = True
form3.Label3.Visible = True
form3.Label4.Visible = True
form3.Label5.Visible = True
form3.Text2.Text = Empty
form3.Adodc1.Enabled = True
form3.Text6.Enabled = True
form3.Text7.Enabled = True
MsgBox "correct password"
Form2.Hide
form3.Show
Text1.Text = Empty
form3.Frame3.Visible = True
form3.Frame1.Visible = True
form3.Frame2.Visible = True
form3.Adodc2.Visible = True
Else
MsgBox "incorrect code"
End If
End Sub
Private Sub Command3_Click()
form3.Frame3.Visible = False
form3.Frame1.Visible = False
form3.Frame2.Visible = False
form3.Adodc2.Visible = False
Form2.Hide
form3.Show
End Sub

Private Sub Label1_Click()
Timer1.Enabled = True
End Sub

Private Sub Label2_Click()
Label2.Left = -5500
Timer1.Enabled = False
Label3.Top = 240
End Sub

Private Sub Label3_Click()
Form2.Hide
Form1.Show
End Sub

Private Sub Timer1_Timer()
Label2.Left = Val(Label2.Left) + 30
Label3.Top = Val(Label3.Top) - 10
If Label2.Left >= 0 Then
Timer1.Enabled = False
End If
End Sub
