VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "log in screen"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   MouseIcon       =   "Form3.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "submit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   840
      TabIndex        =   4
      Top             =   2040
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "#"
      TabIndex        =   1
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "password"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "name"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = Text1.Tag And Text2.Text = Text2.Tag Then
MsgBox "welcome!"
Form3.Hide
Form1.Show
Else
MsgBox "get lost hacker!!!"
End
End If
End Sub
