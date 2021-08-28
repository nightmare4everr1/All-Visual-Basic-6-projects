VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "Form5"
   ClientHeight    =   9855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12165
   FillColor       =   &H00FF8080&
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9855
   ScaleWidth      =   12165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   3600
      Top             =   6480
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   5520
      Top             =   7080
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5040
      Top             =   7080
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "max marks"
      Height          =   495
      Left            =   3960
      TabIndex        =   19
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "canidate score"
      Height          =   495
      Left            =   3960
      TabIndex        =   18
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "roll no"
      Height          =   495
      Left            =   3960
      TabIndex        =   17
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "student name"
      Height          =   495
      Left            =   3960
      TabIndex        =   16
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      DataField       =   "canidatename"
      DataSource      =   "Adodc2"
      Height          =   495
      Left            =   5280
      TabIndex        =   15
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      DataField       =   "canidaterol no"
      DataSource      =   "Adodc2"
      Height          =   495
      Left            =   5280
      TabIndex        =   14
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      DataField       =   "canidatescore"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   5280
      TabIndex        =   13
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      DataField       =   "totalmarks"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   5280
      TabIndex        =   12
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "your marvelous result has earned you this prestigous diploma!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   2640
      TabIndex        =   11
      Top             =   2160
      Visible         =   0   'False
      Width           =   7335
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   4440
      TabIndex        =   10
      Top             =   8520
      Width           =   255
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   6960
      TabIndex        =   9
      Top             =   8280
      Width           =   2775
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   6240
      TabIndex        =   8
      Top             =   8280
      Width           =   735
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   8160
      TabIndex        =   7
      Top             =   7800
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "which gives you a grade of"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6240
      TabIndex        =   6
      Top             =   7800
      Width           =   2055
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "From saint mary's academy"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2640
      TabIndex        =   5
      Top             =   7680
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   3960
      TabIndex        =   4
      Top             =   8520
      Width           =   375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "You have achieved a result percentage of"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   8280
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "MS SystemEx"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Print Layout"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   5040
      TabIndex        =   1
      Top             =   1080
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Result"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   855
      Left            =   5400
      TabIndex        =   0
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Line Line4 
      X1              =   2400
      X2              =   2400
      Y1              =   1440
      Y2              =   9360
   End
   Begin VB.Line Line3 
      X1              =   10080
      X2              =   2400
      Y1              =   9360
      Y2              =   9360
   End
   Begin VB.Line Line2 
      X1              =   10080
      X2              =   10080
      Y1              =   1440
      Y2              =   9360
   End
   Begin VB.Line Line1 
      X1              =   2400
      X2              =   10080
      Y1              =   1440
      Y2              =   1440
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
End
End Sub

Private Sub Form_Activate()
Label16.Caption = Form4.Label5.Caption
Label15.Caption = Form4.Label7.Caption
Label14.Caption = Form4.Label8.Caption
Label13.Caption = Form4.Label9.Caption
End Sub

Private Sub Label15_Click()

End Sub

Private Sub Label3_Click()
Form5.Hide
Form4.Show
End Sub

Private Sub Timer1_Timer()
Label5.Caption = Val(Form4.Label8.Caption) / Val(Form4.Label9.Caption) * 100

End Sub

Private Sub Timer2_Timer()
Dim mark As Single
mark = Label5.Caption
Select Case mark
Case Is >= 90
Label10.Caption = "your work is admirable!congrats!!"
Label8.Caption = "A"
Label12.Visible = True
Case Is >= 80
Label10.Caption = "very good!"
Label8.Caption = "B"
Case Is >= 70
Label10.Caption = "good"
Label8.Caption = "C"
Case Is >= 60
Label10.Caption = "you need to improve"
Label8.Caption = "D"
Case Is >= 50
Label10.Caption = "your result is unsatisfactory"
Label8.Caption = "E"
Case Is >= 40
Label10.Caption = "you have failed the test"
Label8.Caption = "F"
Case Is >= 0
Label10.Caption = "(ungraded)no comments...."
Label8.Caption = "U"
End Select
End Sub

Private Sub Timer3_Timer()
If Label5.Caption < 90 Then
Label12.Visible = False
End If
End Sub
