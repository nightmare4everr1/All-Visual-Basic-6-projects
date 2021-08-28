VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13515
   LinkTopic       =   "Form1"
   ScaleHeight     =   9915
   ScaleWidth      =   13515
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer4 
      Interval        =   1
      Left            =   2040
      Top             =   7800
   End
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   1320
      Top             =   7800
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   720
      Top             =   7800
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   7800
   End
   Begin VB.CommandButton Command6 
      Caption         =   "enter value for label4"
      Height          =   615
      Left            =   3960
      TabIndex        =   14
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "enter value for label3"
      Height          =   615
      Left            =   3960
      TabIndex        =   13
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "enter value for label2"
      Height          =   615
      Left            =   3960
      TabIndex        =   12
      Top             =   5280
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   2055
      Left            =   5520
      TabIndex        =   11
      Top             =   4680
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "enter value for label1"
      Height          =   615
      Left            =   3960
      TabIndex        =   10
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "sort"
      Height          =   615
      Left            =   1200
      TabIndex        =   9
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "clear label5,6,7,8"
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      Top             =   6720
      Width           =   1455
   End
   Begin VB.Line Line1 
      X1              =   600
      X2              =   2280
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   495
      Left            =   5280
      TabIndex        =   8
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   495
      Left            =   5280
      TabIndex        =   7
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "4"
      Height          =   495
      Left            =   840
      TabIndex        =   6
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   495
      Left            =   5280
      TabIndex        =   5
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   495
      Left            =   5280
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "1"
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "2"
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "3"
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   2040
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Label5.Caption = Empty
Label6.Caption = Empty
Label7.Caption = Empty
Label8.Caption = Empty
End Sub

Private Sub Command2_Click()
If Val(Label1.Caption) > Val(Label2.Caption) And Val(Label1.Caption) > Val(Label3.Caption) And Val(Label1.Caption) > Val(Label4.Caption) Then
Label5.Caption = Label1.Caption
End If
If Val(Label2.Caption) < Val(Label1.Caption) And Val(Label2.Caption) > Val(Label3.Caption) And Val(Label2.Caption) > Val(Label4.Caption) Then
Label6.Caption = Label2.Caption
End If
If Val(Label3.Caption) < Val(Label1.Caption) And Val(Label3.Caption) < Val(Label2.Caption) And Val(Label3.Caption) > Val(Label4.Caption) Then
Label7.Caption = Label3.Caption
End If
If Val(Label4.Caption) < Val(Label1.Caption) And Val(Label4.Caption) < Val(Label3.Caption) And Val(Label4.Caption) < Val(Label2.Caption) Then
Label8.Caption = Label4.Caption
End If
'____________________________________________________________________________________________________________________________________________'

If Val(Label2.Caption) > Val(Label1.Caption) And Val(Label2.Caption) > Val(Label3.Caption) And Val(Label2.Caption) > Val(Label4.Caption) Then
Label5.Caption = Label2.Caption
End If
If Val(Label1.Caption) < Val(Label2.Caption) And Val(Label1.Caption) > Val(Label3.Caption) And Val(Label1.Caption) > Val(Label4.Caption) Then
Label6.Caption = Label1.Caption
End If
If Val(Label3.Caption) < Val(Label2.Caption) And Val(Label3.Caption) < Val(Label1.Caption) And Val(Label3.Caption) > Val(Label4.Caption) Then
Label7.Caption = Label3.Caption
End If
If Val(Label4.Caption) < Val(Label1.Caption) And Val(Label4.Caption) < Val(Label3.Caption) And Val(Label4.Caption) < Val(Label2.Caption) Then
Label8.Caption = Label4.Caption
'____________________________________________________________________________________________________________________________________________'

End If
If Val(Label3.Caption) > Val(Label1.Caption) And Val(Label3.Caption) > Val(Label2.Caption) And Val(Label3.Caption) > Val(Label4.Caption) Then
Label5.Caption = Label3.Caption
End If
If Val(Label1.Caption) < Val(Label3.Caption) And Val(Label1.Caption) > Val(Label2.Caption) And Val(Label1.Caption) > Val(Label4.Caption) Then
Label6.Caption = Label1.Caption
End If
If Val(Label2.Caption) < Val(Label3.Caption) And Val(Label2.Caption) < Val(Label1.Caption) And Val(Label2.Caption) > Val(Label4.Caption) Then
Label7.Caption = Label2.Caption
End If
If Val(Label4.Caption) < Val(Label1.Caption) And Val(Label4.Caption) < Val(Label3.Caption) And Val(Label4.Caption) < Val(Label2.Caption) Then
Label8.Caption = Label4.Caption
End If
'____________________________________________________________________________________________________________________________________________'
If Val(Label4.Caption) > Val(Label1.Caption) And Val(Label4.Caption) > Val(Label3.Caption) And Val(Label4.Caption) > Val(Label2.Caption) Then
Label5.Caption = Label4.Caption
If Val(Label1.Caption) < Val(Label4.Caption) And Val(Label1.Caption) > Val(Label2.Caption) And Val(Label1.Caption) > Val(Label3.Caption) Then
Label6.Caption = Label1.Caption
End If
If Val(Label2.Caption) < Val(Label4.Caption) And Val(Label2.Caption) < Val(Label1.Caption) And Val(Label2.Caption) > Val(Label3.Caption) Then
Label7.Caption = Label2.Caption
End If
If Val(Label3.Caption) < Val(Label4.Caption) And Val(Label3.Caption) < Val(Label1.Caption) And Val(Label3.Caption) < Val(Label2.Caption) Then
Label8.Caption = Label3.Caption
End If
End If
End Sub

Private Sub Command3_Click()
Label1.Caption = Text1.Text
End Sub

Private Sub Command4_Click()
Label2.Caption = Text1.Text
End Sub

Private Sub Command5_Click()
Label3.Caption = Text1.Text
End Sub

Private Sub Command6_Click()
Label4.Caption = Text1.Text
End Sub

Private Sub Text1_Click()
Text1.Text = Empty
End Sub

Private Sub Timer1_Timer()
If Val(Label1.Caption) > Val(Label2.Caption) And Val(Label1.Caption) > Val(Label3.Caption) And Val(Label1.Caption) > Val(Label4.Caption) Then
Label5.Caption = Label1.Caption
End If
If Val(Label2.Caption) < Val(Label1.Caption) And Val(Label2.Caption) > Val(Label3.Caption) And Val(Label2.Caption) > Val(Label4.Caption) Then
Label6.Caption = Label2.Caption
End If
If Val(Label3.Caption) < Val(Label1.Caption) And Val(Label3.Caption) < Val(Label2.Caption) And Val(Label3.Caption) > Val(Label4.Caption) Then
Label7.Caption = Label3.Caption
End If
If Val(Label4.Caption) < Val(Label1.Caption) And Val(Label4.Caption) < Val(Label3.Caption) And Val(Label4.Caption) < Val(Label2.Caption) Then
Label8.Caption = Label4.Caption
End If
End Sub

Private Sub Timer2_Timer()
If Val(Label2.Caption) > Val(Label1.Caption) And Val(Label2.Caption) > Val(Label3.Caption) And Val(Label2.Caption) > Val(Label4.Caption) Then
Label5.Caption = Label2.Caption
End If
If Val(Label1.Caption) < Val(Label2.Caption) And Val(Label1.Caption) > Val(Label3.Caption) And Val(Label1.Caption) > Val(Label4.Caption) Then
Label6.Caption = Label1.Caption
End If
If Val(Label3.Caption) < Val(Label2.Caption) And Val(Label3.Caption) < Val(Label1.Caption) And Val(Label3.Caption) > Val(Label4.Caption) Then
Label7.Caption = Label3.Caption
End If
If Val(Label4.Caption) < Val(Label1.Caption) And Val(Label4.Caption) < Val(Label3.Caption) And Val(Label4.Caption) < Val(Label2.Caption) Then
Label8.Caption = Label4.Caption
End If
End Sub

Private Sub Timer3_Timer()
If Val(Label3.Caption) > Val(Label1.Caption) And Val(Label3.Caption) > Val(Label2.Caption) And Val(Label3.Caption) > Val(Label4.Caption) Then
Label5.Caption = Label3.Caption
End If
If Val(Label1.Caption) < Val(Label3.Caption) And Val(Label1.Caption) > Val(Label2.Caption) And Val(Label1.Caption) > Val(Label4.Caption) Then
Label6.Caption = Label1.Caption
End If
If Val(Label2.Caption) < Val(Label3.Caption) And Val(Label2.Caption) < Val(Label1.Caption) And Val(Label2.Caption) > Val(Label4.Caption) Then
Label7.Caption = Label2.Caption
End If
If Val(Label4.Caption) < Val(Label1.Caption) And Val(Label4.Caption) < Val(Label3.Caption) And Val(Label4.Caption) < Val(Label2.Caption) Then
Label8.Caption = Label4.Caption
End If
End Sub

Private Sub Timer4_Timer()
If Val(Label4.Caption) > Val(Label1.Caption) And Val(Label4.Caption) > Val(Label3.Caption) And Val(Label4.Caption) > Val(Label2.Caption) Then
Label5.Caption = Label4.Caption
If Val(Label1.Caption) < Val(Label4.Caption) And Val(Label1.Caption) > Val(Label2.Caption) And Val(Label1.Caption) > Val(Label3.Caption) Then
Label6.Caption = Label1.Caption
End If
If Val(Label2.Caption) < Val(Label4.Caption) And Val(Label2.Caption) < Val(Label1.Caption) And Val(Label2.Caption) > Val(Label3.Caption) Then
Label7.Caption = Label2.Caption
End If
If Val(Label3.Caption) < Val(Label4.Caption) And Val(Label3.Caption) < Val(Label1.Caption) And Val(Label3.Caption) < Val(Label2.Caption) Then
Label8.Caption = Label3.Caption
End If
End If
End Sub
