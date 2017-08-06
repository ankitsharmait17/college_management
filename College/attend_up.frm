VERSION 5.00
Begin VB.Form Form21 
   Caption         =   "Form21"
   ClientHeight    =   9150
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15195
   LinkTopic       =   "Form21"
   ScaleHeight     =   9150
   ScaleWidth      =   15195
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Height          =   735
      Left            =   0
      Picture         =   "attend_up.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Submit"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8040
      Width           =   2295
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00FFFF80&
      Height          =   480
      Left            =   6600
      TabIndex        =   6
      Top             =   6720
      Width           =   4695
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00FFFF80&
      Height          =   480
      Left            =   6600
      TabIndex        =   5
      Top             =   5960
      Width           =   4695
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFF80&
      Height          =   480
      Left            =   6600
      TabIndex        =   4
      Top             =   5200
      Width           =   4695
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFF80&
      Height          =   480
      Left            =   6600
      TabIndex        =   3
      Top             =   4440
      Width           =   4695
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFF80&
      Height          =   480
      Left            =   6600
      TabIndex        =   2
      Top             =   3680
      Width           =   4695
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFF80&
      Height          =   480
      Left            =   6600
      TabIndex        =   1
      Top             =   2920
      Width           =   4695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   6600
      TabIndex        =   0
      Top             =   2160
      Width           =   4695
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Update Student Attendance"
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
      Height          =   735
      Left            =   1560
      TabIndex        =   16
      Top             =   240
      Width           =   9975
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFC0&
      Caption         =   "subject code"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   1440
      TabIndex        =   15
      Top             =   5200
      Width           =   4695
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFC0&
      Caption         =   "last employer"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   1440
      TabIndex        =   14
      Top             =   6720
      Width           =   4695
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   1440
      TabIndex        =   13
      Top             =   5960
      Width           =   4695
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "course code"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   1440
      TabIndex        =   12
      Top             =   4440
      Width           =   4695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "date"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   1440
      TabIndex        =   11
      Top             =   3680
      Width           =   4695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "roll"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   1440
      TabIndex        =   10
      Top             =   2920
      Width           =   4695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "attd_ id"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Top             =   2160
      Width           =   4695
   End
End
Attribute VB_Name = "Form21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
