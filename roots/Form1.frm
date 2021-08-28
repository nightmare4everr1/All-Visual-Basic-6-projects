VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Under root system"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      Caption         =   "Terminate program"
      Height          =   495
      Left            =   3480
      TabIndex        =   13
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Hide process"
      Height          =   615
      Left            =   1560
      TabIndex        =   12
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Show process"
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5880
      Top             =   3840
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5280
      Top             =   3840
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4800
      Top             =   3840
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4320
      Top             =   3840
   End
   Begin VB.CommandButton Command3 
      Caption         =   "END"
      Height          =   615
      Left            =   5280
      TabIndex        =   8
      Top             =   4440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Find roots(high accuracy)"
      Height          =   615
      Left            =   360
      TabIndex        =   7
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   735
      Left            =   1560
      TabIndex        =   3
      Top             =   3480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   3480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3840
      Top             =   3840
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Find roots(standard accuracy)"
      Height          =   735
      Left            =   360
      TabIndex        =   1
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   5040
      TabIndex        =   0
      Text            =   "0"
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Do not alter the values"
      Height          =   255
      Left            =   600
      TabIndex        =   11
      Top             =   3120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Text4 
      Height          =   495
      Left            =   5040
      TabIndex        =   9
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "To find roots of numbers from 0 to 9.9000000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1575
      Left            =   960
      TabIndex        =   6
      Top             =   0
      Width           =   5295
   End
   Begin VB.Label Label2 
      Caption         =   "its roots"
      Height          =   495
      Left            =   3960
      TabIndex        =   5
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "integar"
      Height          =   735
      Left            =   3960
      TabIndex        =   4
      Top             =   1560
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = 0 Then
MsgBox "Syntax error", , "Error"
Else
Text2.Text = Empty
Text3.Text = Empty
Text4.Caption = Empty
Timer1.Enabled = True
Command2.Enabled = False
End If
End Sub

Private Sub Command2_Click()
Command3.Visible = True
MsgBox "This will take alot of time,you can click on 'END'to discontinue.the roots will have an accuracy of 0.001.", , "Advice"
If Text1.Text = 0 Then
MsgBox "Syntax error", , "Error"
Else
Text2.Text = Empty
Text3.Text = Empty
Text4.Caption = Empty
Timer2.Enabled = True
Timer3.Enabled = True
Timer4.Enabled = True
Timer5.Enabled = True
Command1.Enabled = False
End If
End Sub

Private Sub Command3_Click()
Text2.Text = Empty
Text3.Text = Empty
Command3.Visible = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
End Sub

Private Sub Command4_Click()
Dim userMsg As String
userMsg = InputBox("Input password", "Password", "Enter your password here", 500, 700)
If userMsg = "kai" Then
MsgBox "Correct password"
Text2.Visible = True
Text3.Visible = True
Label4.Visible = True
Command5.Visible = True
Else
MsgBox "Error! Error ! Shutting down"
End
End If
End Sub

Private Sub Command5_Click()
Text2.Visible = False
Text3.Visible = False
Label4.Visible = False
Command5.Visible = False
End Sub

Private Sub Command6_Click()
End
End Sub

Private Sub Timer1_Timer()
Text2.Text = Val(Text2.Text) + 0.1
Text3.Text = Val(Text2.Text) ^ 2
If Text3.Text >= Text1.Text Then
Text4.Caption = Text2.Text
Timer1.Enabled = False
Command2.Enabled = True
End If
End Sub

Private Sub Timer2_Timer()
Text2.Text = Val(Text2.Text) + 0.001
Text3.Text = Val(Text2.Text) ^ 2
If Text3.Text >= Text1.Text Then
Text4.Caption = Text2.Text
Timer2.Enabled = False
Command1.Enabled = True
Text4.Caption = Val(Text4.Caption) - 0.003
End If
End Sub

Private Sub Timer3_Timer()
Text2.Text = Val(Text2.Text) + 0.001
Text3.Text = Val(Text2.Text) ^ 2
If Text3.Text >= Text1.Text Then
Text4.Caption = Text2.Text
Timer3.Enabled = False
End If
End Sub

Private Sub Timer4_Timer()
Text2.Text = Val(Text2.Text) + 0.001
Text3.Text = Val(Text2.Text) ^ 2
If Text3.Text >= Text1.Text Then
Text4.Caption = Text2.Text
Timer4.Enabled = False
End If
End Sub

Private Sub Timer5_Timer()
Text2.Text = Val(Text2.Text) + 0.001
Text3.Text = Val(Text2.Text) ^ 2
If Text3.Text >= Text1.Text Then
Text4.Caption = Text2.Text
Timer5.Enabled = False
End If
End Sub
