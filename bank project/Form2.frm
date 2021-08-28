VERSION 5.00
Begin VB.Form Form2 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "transaction  receipt"
   ClientHeight    =   9705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13890
   LinkTopic       =   "Form2"
   ScaleHeight     =   9705
   ScaleWidth      =   13890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   960
      Top             =   6360
   End
   Begin VB.ListBox List2 
      Height          =   1035
      ItemData        =   "Form2.frx":0000
      Left            =   4200
      List            =   "Form2.frx":0002
      TabIndex        =   5
      Top             =   3480
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   1035
      ItemData        =   "Form2.frx":0004
      Left            =   1920
      List            =   "Form2.frx":0006
      TabIndex        =   4
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   3600
      Top             =   360
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6360
      Top             =   240
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   7
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "back"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   24
         Charset         =   1
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   6
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "deposits"
      BeginProperty Font 
         Name            =   "Marking Pen"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   4200
      TabIndex        =   3
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "withdrawels"
      BeginProperty Font 
         Name            =   "Marking Pen"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "your transaction history"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   2040
      Width           =   4815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "name"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4320
      TabIndex        =   0
      Top             =   1200
      Width           =   4215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If List1.ListCount > 11 Then
List1.RemoveItem (0)
Form1.Adodc1.Recordset.Fields(8) = Val(Form1.Adodc1.Recordset.Fields(8)) - 1
Form1.Adodc1.Recordset.Save
End If
End Sub

Private Sub Command2_Click()
Open "C:\Microsoft Visual Studio\VB98\bank project\withdraw/a.Txt" For Input As #1
Delete# 1
End Sub

Private Sub Command3_Click()
Form1.Show
End Sub

Private Sub Form_Load()
Label1.Caption = Form1.Adodc1.Recordset.Fields(0)
End Sub

Private Sub Label5_Click()
Form1.Show
Form2.Hide
End Sub

Private Sub Timer1_Timer()
On Error GoTo err:
X = Form1.Adodc1.Recordset.Fields(0)
Y = 0
Do While Y <= X
Y = Y + 1
Dim variable1 As String
Dim a As Integer
Dim b As String
For a = 0 To 30
b = Form1.Adodc1.Recordset.Fields(2)
Open "C:\Microsoft Visual Studio\VB98\bank project\withdraw\" & a & b & ".Txt" For Input As #1
Input #1, variable1
List1.AddItem (variable1)
Close #1
Next a
Loop
Exit Sub
err:
Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
On Error GoTo err:
X = Form1.Adodc1.Recordset.Fields(0)
Y = 0
Do While Y <= X
Y = Y + 1
Dim variable2 As String
Dim a As Integer
Dim b As String
For a = 0 To 30
b = Form1.Adodc1.Recordset.Fields(2)
Open "C:\Microsoft Visual Studio\VB98\bank project\deposit\" & a & b & ".Txt" For Input As #1
Input #1, variable2
List2.AddItem (variable2)
Close #1
Next a
Loop
Exit Sub
err:
Timer2.Enabled = False
Timer1.Enabled = True
End Sub

Private Sub Timer3_Timer()
If List1.ListCount > 11 Then
List1.RemoveItem (0)
Form1.Adodc1.Recordset.Fields(8) = Val(Form1.Adodc1.Recordset.Fields(8)) - 1
Dim a As Integer
Dim b As String
a = Form1.Adodc1.Recordset.Fields(8)
b = Form1.Adodc1.Recordset.Fields(2)
Unload "C:\Microsoft Visual Studio\VB98\bank project\withdraw/" & a & b & ".Txt"
Form1.Adodc1.Recordset.Save
End If
End Sub
