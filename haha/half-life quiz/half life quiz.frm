VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "half-life2 quiz"
   ClientHeight    =   10530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9915
   LinkTopic       =   "Form1"
   ScaleHeight     =   10530
   ScaleWidth      =   9915
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   4560
      Top             =   7200
   End
   Begin VB.CommandButton Command6 
      Caption         =   "click here to forward on to the pokemon quiz"
      Height          =   735
      Left            =   3240
      TabIndex        =   22
      Top             =   8520
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "submit result"
      Height          =   735
      Left            =   7200
      TabIndex        =   21
      Top             =   7680
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   3240
      TabIndex        =   16
      Top             =   8040
      Width           =   3495
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   2280
      TabIndex        =   13
      Top             =   8040
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Height          =   735
      Left            =   8640
      TabIndex        =   12
      Top             =   9480
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   735
      Left            =   2640
      TabIndex        =   11
      Top             =   9480
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   8400
      TabIndex        =   10
      Top             =   6480
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   2640
      TabIndex        =   9
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "submit"
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   9840
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "submit"
      Height          =   375
      Left            =   5640
      TabIndex        =   7
      Top             =   6840
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "submit"
      Height          =   375
      Left            =   5640
      TabIndex        =   6
      Top             =   9840
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "submit"
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   6960
      Width           =   2055
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      ItemData        =   "half life quiz.frx":0000
      Left            =   5640
      List            =   "half life quiz.frx":0013
      TabIndex        =   4
      Top             =   9480
      Width           =   2655
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "half life quiz.frx":0061
      Left            =   360
      List            =   "half life quiz.frx":0071
      TabIndex        =   3
      Top             =   9480
      Width           =   2055
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "half life quiz.frx":00A5
      Left            =   5640
      List            =   "half life quiz.frx":00B5
      TabIndex        =   2
      Top             =   6480
      Width           =   2535
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "half life quiz.frx":00FA
      Left            =   480
      List            =   "half life quiz.frx":010A
      TabIndex        =   1
      Top             =   6480
      Width           =   2055
   End
   Begin VB.PictureBox Picture2 
      Height          =   5535
      Left            =   360
      Picture         =   "half life quiz.frx":0139
      ScaleHeight     =   5475
      ScaleWidth      =   9075
      TabIndex        =   0
      Top             =   120
      Width           =   9135
   End
   Begin VB.Label Label6 
      Caption         =   "this creature has a bullet proof armour how can we defeat it"
      Height          =   615
      Left            =   360
      TabIndex        =   20
      Top             =   8760
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "what is its weapon"
      Height          =   495
      Left            =   5520
      TabIndex        =   19
      Top             =   5880
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "what is the name of the gun  infront of you"
      Height          =   495
      Left            =   5640
      TabIndex        =   18
      Top             =   8760
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "what is this creature"
      Height          =   495
      Left            =   480
      TabIndex        =   17
      Top             =   5880
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "rank"
      Height          =   255
      Left            =   3360
      TabIndex        =   15
      Top             =   7560
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "total score"
      Height          =   255
      Left            =   2280
      TabIndex        =   14
      Top             =   7560
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label7_Click()
End Sub

Private Sub Command1_Click()
If Combo1 = "strider" Then
Text1 = "2"
MsgBox "correct"
Combo1.Enabled = False
Else
MsgBox "wrong"
Combo1.Enabled = False
Text1 = "0"
End If
End Sub

Private Sub Command2_Click()
If Combo4 = "pulse rifle" Then
Text4 = "3"
MsgBox "correct"
Combo4.Enabled = False
Command5.Visible = True
Else
MsgBox "wrong"
Combo4.Enabled = False
Text4 = "0"
Command5.Visible = True
End If
End Sub

Private Sub Command3_Click()
If Combo2 = "particle disintegrator" Then
Text2 = "2"
MsgBox "correct"
Combo2.Enabled = False
Else
MsgBox "wrong"
Combo2.Enabled = False
Text2 = "0"
End If
End Sub

Private Sub Command4_Click()
If Combo3 = "any explosives" Then
Text3 = "1"
MsgBox "correct"
Combo3.Enabled = False
Else
MsgBox "wrong"
Combo3.Enabled = False
Text3 = "0"
End If
End Sub

Private Sub Command5_Click()
Text5 = Val(Text1.Text) + Val(Text2.Text) + Val(Text3.Text) + Val(Text4.Text)
Timer1.Enabled = True
If Val(Text1.Text) + Val(Text2.Text) + Val(Text3.Text) + Val(Text4.Text) = 8 Then
Text6.Text = "your a great fan"
Else
If Val(Text1.Text) + Val(Text2.Text) + Val(Text3.Text) + Val(Text4.Text) = 7 Then
Text6.Text = "your a great fan"
Else
If Val(Text1.Text) + Val(Text2.Text) + Val(Text3.Text) + Val(Text4.Text) = 6 Then
Text6.Text = "pretty good"
Else
If Val(Text1.Text) + Val(Text2.Text) + Val(Text3.Text) + Val(Text4.Text) = 5 Then
Text6.Text = "pretty good"
Else
If Val(Text1.Text) + Val(Text2.Text) + Val(Text3.Text) + Val(Text4.Text) = 4 Then
Text6.Text = "not good"
Else
If Val(Text1.Text) + Val(Text2.Text) + Val(Text3.Text) + Val(Text4.Text) = 3 Then
Text6.Text = "lame"
Else
If Val(Text1.Text) + Val(Text2.Text) + Val(Text3.Text) + Val(Text4.Text) = 2 Then
Text6.Text = "lame"
Else
If Val(Text1.Text) + Val(Text2.Text) + Val(Text3.Text) + Val(Text4.Text) = 1 Then
Text6.Text = "phetatic, tucka nahi chalta"
Else
If Val(Text1.Text) + Val(Text2.Text) + Val(Text3.Text) + Val(Text4.Text) = 0 Then
Text6.Text = "what a phetatic loser!"
End If
End If
End If
End If
End If
End If
End If
End If
End If
End Sub

Private Sub Command6_Click()
Form1.Hide
Form2.Show
End Sub

Private Sub Form_Load()
MsgBox "please verify password"
Form1.Hide
Form3.Show
End Sub

Private Sub Timer1_Timer()
Command6.Visible = True
End Sub
