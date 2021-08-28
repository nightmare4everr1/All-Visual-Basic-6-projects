VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   840
      TabIndex        =   1
      Tag             =   "fahad"
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "submit"
      Height          =   735
      Left            =   840
      TabIndex        =   0
      Top             =   600
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
Form2.Hide
Form1.Show
MsgBox "welcome"
Else
Form1.Adodc1.Enabled = False
Form2.Hide
Form1.Show
MsgBox "welcome"
Form1.Enabled = False
End If
End Sub

