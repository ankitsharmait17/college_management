VERSION 5.00
Begin VB.Form Form20 
   Caption         =   "Form20"
   ClientHeight    =   8790
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13935
   LinkTopic       =   "Form20"
   ScaleHeight     =   8790
   ScaleWidth      =   13935
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text7 
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
      TabIndex        =   15
      Top             =   1440
      Width           =   4695
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Height          =   735
      Left            =   0
      Picture         =   "Form20.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
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
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7200
      Width           =   2295
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
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Mark Id"
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
      TabIndex        =   16
      Top             =   1440
      Width           =   4695
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Marks"
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
      Left            =   5160
      TabIndex        =   14
      Top             =   360
      Width           =   5295
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFC0&
      Caption         =   "year"
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
      Top             =   5200
      Width           =   4695
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "semester"
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
      Top             =   5960
      Width           =   4695
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "grade"
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
      Top             =   4440
      Width           =   4695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "marks"
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
      Top             =   3680
      Width           =   4695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Subject Code"
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
      Top             =   2920
      Width           =   4695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Roll no."
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
      TabIndex        =   8
      Top             =   2160
      Width           =   4695
   End
End
Attribute VB_Name = "Form20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim cmd As New ADODB.Command
Dim rst As New ADODB.Recordset
Private Sub Command1_Click()
With cmd
.ActiveConnection = conn
.CommandText = "Insert into Marks values('" & Text7.Text & "','" & Text1.Text & "','" & Text2.Text & "','" & Text3.Text & "','" & Text4.Text & "','" & Text5.Text & "','" & Text6.Text & "')"
.CommandType = adCmdText
conn.BeginTrans 'to insert a new row
.Execute 'to insert the data
conn.CommitTrans 'to save the data


MsgBox ("data inserted successfully")
.CommandText = "select * from marks"
rst.Close
rst.Open cmd
End With
End Sub

Private Sub Command2_Click()
Unload Me
Load Form3
Form3.Show
End Sub

Private Sub Form_Load()

conn.ConnectionString = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\" & "vb1.mdb; Mode= Read|Write"
    conn.CursorLocation = adUseClient
    conn.Open

With cmd
.ActiveConnection = conn
  .CommandText = "SELECT * From marks where roll='" & Text1.Text & "'"
.CommandType = adCmdText
  End With

With rst
    .CursorType = adOpenStatic
    .CursorLocation = adUseClient
    .LockType = adLockOptimistic
    .Open cmd
End With
Text1.Text = Form19.Text1.Text
Text2.Text = Form19.Text2.Text

End Sub
