VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Welcome ADMIN"
   ClientHeight    =   9990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14460
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   9990
   ScaleWidth      =   14460
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command9 
      BackColor       =   &H00FFFFC0&
      Caption         =   "OK"
      Height          =   735
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FFFFC0&
      Caption         =   "OK"
      Height          =   735
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.OptionButton Option7 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Faculty"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   13
      Top             =   5400
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.OptionButton Option6 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Student"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   12
      Top             =   5400
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.OptionButton Option5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Faculty"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7920
      TabIndex        =   11
      Top             =   5400
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Student"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11400
      TabIndex        =   10
      Top             =   5400
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Display Members"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3840
      Width           =   4935
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Delete Member"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3720
      Width           =   5415
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFC0C0&
      Height          =   735
      Left            =   120
      Picture         =   "Form2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9120
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "OK"
      Height          =   735
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "B.Tech"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9960
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "OK"
      Height          =   735
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Student"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      TabIndex        =   3
      Top             =   1440
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Faculty"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1440
      TabIndex        =   2
      Top             =   1440
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Update Course"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   5415
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Create an account"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   5415
   End
   Begin VB.Image Image1 
      Height          =   9690
      Left            =   0
      Picture         =   "Form2.frx":059A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14340
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer
Private Sub Command1_Click()
Option1.Enabled = True
Option1.Visible = True
Option2.Enabled = True
Option2.Visible = True
Command3.Enabled = True
Command3.Visible = True

End Sub

Private Sub Command2_Click()
Option3.Enabled = True
Option3.Visible = True
Command4.Enabled = True
Command4.Visible = True

End Sub

Private Sub Command3_Click()
If Option1.Value = True Then
Unload Me
Load Form9
Form9.Show
ElseIf Option2.Value = True Then
Unload Me
Load Form10
Form10.Show
Else
MsgBox ("Enter a choice!")
End If




End Sub



Private Sub Command4_Click()

If Option3.Value = True Then
Unload Me
Load Form12
Form12.Show
ElseIf Option4.Value = True Then
Unload Me
Load Form13
Form13.Show
Else
MsgBox ("Enter a choice!")
End If




End Sub

Private Sub Command5_Click()
'If rst.State <> adStateClosed Then rst.Close
'If conn.State <> adStateClosed Then conn.Close

Unload Me
Unload Form1
Load Form1
Form1.Show
End Sub

Private Sub Command6_Click()
Option4.Enabled = True
Option4.Visible = True
Option5.Enabled = True
Option5.Visible = True
Command8.Enabled = True
Command8.Visible = True

End Sub

Private Sub Command7_Click()
Option7.Enabled = True
Option7.Visible = True
Option6.Enabled = True
Option6.Visible = True
Command9.Enabled = True
Command9.Visible = True

End Sub

Private Sub Command8_Click()
If Option5.Value = True Then
Unload Me
Form14.Show
ElseIf Option4.Value = True Then
Unload Me
Form15.Show
Else
MsgBox ("Enter a choice!")
End If
End Sub

Private Sub Command9_Click()
If Option7.Value = True Then
Unload Me
Form13.Show
ElseIf Option6.Value = True Then
Unload Me
Form16.Show
Else
MsgBox ("Enter a choice!")
End If
End Sub

