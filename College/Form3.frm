VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Form3"
   ClientHeight    =   9990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14460
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   9990
   ScaleWidth      =   14460
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Log Out"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8760
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFF80&
      Caption         =   "Update Attendance"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3480
      Width           =   4215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFF80&
      Caption         =   "Insert Marks"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5040
      Width           =   4215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFF80&
      Caption         =   "change password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6840
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Update Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   4215
   End
   Begin VB.Image Image1 
      Height          =   2595
      Left            =   480
      Picture         =   "Form3.frx":0000
      Stretch         =   -1  'True
      Top             =   360
      Width           =   2460
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Faculty Profile"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1215
      Left            =   4680
      TabIndex        =   0
      Top             =   240
      Width           =   4455
   End
   Begin VB.Image Image2 
      Height          =   9405
      Left            =   0
      Picture         =   "Form3.frx":1088
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15900
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
Load facup
facup.Show

End Sub

Private Sub Command2_Click()
Unload Me
Load change
change.Show
End Sub

Private Sub Command3_Click()
Unload Me
Load Form19
Form19.Show
End Sub

Private Sub Command4_Click()
Load Form18
Form18.Show
End Sub

Private Sub Command5_Click()
Unload Me
Load Form5
Form5.Show

End Sub

Private Sub Command6_Click()

End Sub

