VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9660
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12405
   LinkTopic       =   "Form1"
   ScaleHeight     =   9660
   ScaleWidth      =   12405
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4080
      TabIndex        =   16
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "add"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   15
      Top             =   4920
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1680
      TabIndex        =   14
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "add"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   13
      Top             =   4920
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "STOP SORTING"
      Height          =   495
      Left            =   2520
      TabIndex        =   12
      Top             =   1200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "move to list1"
      Height          =   495
      Left            =   4320
      TabIndex        =   7
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "move to list2"
      Height          =   495
      Left            =   1080
      TabIndex        =   6
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3360
      Top             =   2520
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   840
      Top             =   840
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   240
      Top             =   840
   End
   Begin VB.ListBox List3 
      Height          =   3765
      ItemData        =   "Form1.frx":0000
      Left            =   2400
      List            =   "Form1.frx":0002
      TabIndex        =   3
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "sort into one list"
      Height          =   735
      Left            =   2520
      TabIndex        =   2
      Top             =   1680
      Width           =   1455
   End
   Begin VB.ListBox List2 
      Height          =   1425
      ItemData        =   "Form1.frx":0004
      Left            =   3600
      List            =   "Form1.frx":0011
      TabIndex        =   1
      Top             =   3240
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   1425
      ItemData        =   "Form1.frx":0024
      Left            =   1200
      List            =   "Form1.frx":0034
      TabIndex        =   0
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   255
      Left            =   4440
      TabIndex        =   11
      Top             =   1560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   255
      Left            =   1440
      TabIndex        =   10
      Top             =   1560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "List 2"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   3840
      TabIndex        =   9
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "List 1"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   1320
      TabIndex        =   8
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Line Line1 
      X1              =   3240
      X2              =   3240
      Y1              =   1200
      Y2              =   5520
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Left            =   960
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Do While List3.ListCount <> 0
     List3.RemoveItem (0)
    Loop
Timer3.Enabled = True
Command2.Visible = False
Command1.Visible = False
Command3.Visible = False
Command4.Visible = True
Command5.Visible = False
Text1.Visible = False
Command6.Visible = False
Text2.Visible = False
End Sub

Private Sub Command2_Click()
'Timer3.Enabled = True
'Exit Sub
On Error GoTo err:
If List1.ListIndex < Val(Label1.Caption) - 1 Then
List1.ListIndex = List1.ListIndex + 1
A = List1.Text
'MsgBox "" & A
List2.AddItem (List1.Text)
List1.RemoveItem (List1.ListIndex)
Timer2.Enabled = True
Label2.Caption = "0"
Timer1.Enabled = True
Label1.Caption = "0"
End If
Exit Sub
err:
MsgBox "NO more items in list!"
Label5.Caption = "0"


End Sub

Private Sub Command3_Click()
'Timer3.Enabled = True
'Exit Sub
'On Error GoTo err:
If List2.ListIndex < Val(Label2.Caption) - 1 Then
List2.ListIndex = List2.ListIndex + 1
List1.AddItem (List2.Text)
List2.RemoveItem (List2.ListIndex)
Timer1.Enabled = True
Label1.Caption = "0"
Timer2.Enabled = True
Label2.Caption = "0"
End If
Exit Sub
err:
MsgBox "NO more items in list!"
Label6.Caption = "0"
End Sub

Private Sub Command4_Click()
Timer3.Enabled = False
List1.ListIndex = -1
List2.ListIndex = -1
Command1.Visible = True
Command2.Visible = True
Command3.Visible = True
Command5.Visible = True
Text1.Visible = True
Command6.Visible = True
Text2.Visible = True
End Sub

Private Sub Command5_Click()
If Text1 <> Empty Then
List1.AddItem (Text1.Text)
Text1.Text = Empty
End If
Timer2.Enabled = True
Label2.Caption = "0"
Timer1.Enabled = True
Label1.Caption = "0"
End Sub

Private Sub Command6_Click()
If Text2 <> Empty Then
List2.AddItem (Text2.Text)
Text2.Text = Empty
End If
Timer2.Enabled = True
Label2.Caption = "0"
Timer1.Enabled = True
Label1.Caption = "0"
End Sub

Private Sub Timer1_Timer()
On Error GoTo err:
List1.ListIndex = List1.ListIndex + 1
Label1.Caption = Val(Label1.Caption) + 1
Exit Sub
err:
Timer1.Enabled = False
List1.ListIndex = -1
End Sub

Private Sub Timer2_Timer()
On Error GoTo err:
List2.ListIndex = List2.ListIndex + 1
Label2.Caption = Val(Label2.Caption) + 1
Exit Sub
err:
Timer2.Enabled = False
List2.ListIndex = -1
End Sub

Private Sub Timer3_Timer()

If List1.ListIndex < Val(Label1.Caption) - 1 And Label5.Caption <> "0" Then
List1.ListIndex = List1.ListIndex + 1
List3.AddItem (List1.Text)
End If



If List2.ListIndex < Val(Label2.Caption) - 1 And Label6.Caption <> "0" Then
List2.ListIndex = List2.ListIndex + 1
List3.AddItem (List2.Text)
End If
End Sub

Private Sub Timer4_Timer()
List1.ListIndex = -1
List2.ListIndex = -1
Timer4.Enabled = False
End Sub
