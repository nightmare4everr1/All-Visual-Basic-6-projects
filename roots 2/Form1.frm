VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9900
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13485
   LinkTopic       =   "Form1"
   ScaleHeight     =   9900
   ScaleWidth      =   13485
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   3720
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "find under root"
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   1800
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Y = 0.0001
X = 1
Do While Y ^ 2 < Val(Text1.Text)
Y = Y + 0.0001
X = X + 1
If X > 100000 Then
MsgBox "value too large"
Exit Sub
End If
Loop
Label1.Caption = Y
End Sub

