VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Form6"
   ClientHeight    =   9990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14460
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   Picture         =   "Form6.frx":0000
   ScaleHeight     =   9990
   ScaleWidth      =   14460
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFF80&
      Caption         =   "change password"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   18
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6960
      Width           =   5175
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Log Out"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8880
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFF80&
      Caption         =   "ATTENDANCE STATUS"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   18
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5880
      Width           =   5175
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFF80&
      Caption         =   "PROGRESS report"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   18
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4560
      Width           =   5175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF80&
      Caption         =   "UPDATE INFORMATION"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   18
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3480
      Width           =   5175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Student Profile"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1215
      Left            =   4680
      TabIndex        =   4
      Top             =   840
      Width           =   6375
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Load Form11
Form11.Show

End Sub

Private Sub Command2_Click()

Unload Me
Load Form8
Form8.Show
End Sub

Private Sub Command3_Click()
Load Form7
Form7.Show

End Sub

Private Sub Command4_Click()

Unload Me
Unload Form5
Load Form5
Form5.Show
End Sub

Private Sub Command5_Click()
Unload Me
Load change
change.Show
End Sub

Private Sub Form_Load()

'Me.FontName = "Arial"
'Me.FontSize = 26
'Me.FontBold = True
'Me.FontUnderline = True
'Me.Print "Welcome "
End Sub

