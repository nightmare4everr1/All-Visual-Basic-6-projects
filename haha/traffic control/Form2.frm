VERSION 5.00
Begin VB.Form form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "password verification"
   ClientHeight    =   9855
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   13830
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9855
   ScaleWidth      =   13830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   720
      TabIndex        =   4
      Tag             =   "fahad"
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "submit"
      Height          =   615
      Left            =   2760
      TabIndex        =   3
      Top             =   1080
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   135
      Left            =   3000
      TabIndex        =   2
      Top             =   5040
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CommandButton Command1 
      Caption         =   "submit"
      Enabled         =   0   'False
      Height          =   615
      Left            =   2760
      TabIndex        =   1
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   720
      PasswordChar    =   "*"
      TabIndex        =   0
      Tag             =   "blue"
      Top             =   1920
      Width           =   1935
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = Text1.Tag Then
MsgBox "welcome"
MsgBox "for all newcomers;if you have problems and you think something is malfunctioning and/or if you are having problems using this program please click on HELP which will be shown flashing for a second"
Form2.Hide
Form1.Show
Else
MsgBox "NOW SUFFER HACKER!"
Command1.Enabled = False
Text1.Enabled = False
End If
End Sub

Private Sub Command2_Click()
If Text2.Text = Text2.Tag Then
Text1.Enabled = True
Command1.Enabled = True
Else
Text1.Enabled = False
MsgBox "suffer hacker"
End If
End Sub

Private Sub Command3_Click()
Form2.Hide
Form1.Show
End Sub

Private Sub Form_Load()
MsgBox "remember if you enter password wrong then your computer will become stuck,"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Command1.Enabled = False And Text1.Enabled = False Then
MsgBox "suffer"
Option1.Visible = True
End If
End Sub

Private Sub Form_Resize()
Form2.Show
End Sub

Private Sub Form_Terminate()
Form2.Show
End Sub

Private Sub Timer1_Timer()

End Sub

Private Sub Option1_DblClick()
MsgBox "name fahad,password is your favourite colour"
Option1.Enabled = False
Command1.Enabled = True
Text1.Enabled = True
End Sub

