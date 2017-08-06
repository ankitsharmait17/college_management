VERSION 5.00
Begin VB.Form facup 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Form11"
   ClientHeight    =   9030
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13590
   LinkTopic       =   "Form11"
   ScaleHeight     =   9030
   ScaleWidth      =   13590
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Height          =   735
      Left            =   360
      Picture         =   "facup.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   240
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
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8280
      Width           =   2295
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00FFFF80&
      Height          =   480
      Left            =   6840
      TabIndex        =   6
      Top             =   3055
      Width           =   4695
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00FFFF80&
      Height          =   480
      Left            =   6840
      TabIndex        =   5
      Top             =   3815
      Width           =   4695
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFF80&
      Height          =   480
      Left            =   6840
      TabIndex        =   4
      Top             =   4575
      Width           =   4695
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFF80&
      Height          =   480
      Left            =   6840
      TabIndex        =   3
      Top             =   5335
      Width           =   4695
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFF80&
      Height          =   480
      Left            =   6840
      TabIndex        =   2
      Top             =   6095
      Width           =   4695
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFF80&
      Height          =   480
      Left            =   6840
      TabIndex        =   1
      Top             =   6855
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
      Left            =   6840
      TabIndex        =   0
      Top             =   2295
      Width           =   4695
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Update Profile"
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
      Left            =   5520
      TabIndex        =   15
      Top             =   600
      Width           =   5295
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFC0&
      Caption         =   "qualification"
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
      Left            =   1800
      TabIndex        =   13
      Top             =   5440
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
      Left            =   1800
      TabIndex        =   12
      Top             =   6960
      Width           =   4695
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Years of experience"
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
      Left            =   1800
      TabIndex        =   11
      Top             =   6200
      Width           =   4695
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "contact"
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
      Left            =   1800
      TabIndex        =   10
      Top             =   4680
      Width           =   4695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "address"
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
      Left            =   1800
      TabIndex        =   9
      Top             =   3920
      Width           =   4695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "name"
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
      Left            =   1800
      TabIndex        =   8
      Top             =   3160
      Width           =   4695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "user id"
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
      Left            =   1800
      TabIndex        =   7
      Top             =   2400
      Width           =   4695
   End
End
Attribute VB_Name = "facup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim cmd As New ADODB.Command
Dim rst As New ADODB.Recordset

Private Sub Label9_Click()

End Sub

Private Sub Command1_Click()

With cmd
.ActiveConnection = conn
.CommandText = "update Faculty set F_name='" & Text2.Text & "',F_addr='" & Text3.Text & " ',F_cont='" & Text4.Text & "',F_qualif='" & Text5.Text & "',F_exp='" & Text6.Text & "',Last_emplr='" & Text7.Text & "' where Fac_id='" & Text1.Text & "'"
.CommandType = adCmdText
conn.BeginTrans 'to insert a new row
.Execute 'to insert the data
conn.CommitTrans 'to save the data

MsgBox ("record updated successfully")
.CommandText = "select * from Faculty"
rst.Close
    rst.Open cmd
End With

End Sub

Private Sub Command2_Click()
If rst.State <> adStateClosed Then rst.Close
If conn.State <> adStateClosed Then conn.Close

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
.CommandText = "SELECT * from Faculty where Fac_id='" & Form1.ui.Text & "'"
.CommandType = adCmdText
Form1.ui.Text = ""
Form1.pw.Text = ""
End With
With rst

.CursorType = adOpenStatic
.CursorLocation = adUseClient
.LockType = adLockOptimistic
.Open cmd
End With


Text1.Text = rst.Fields(0)
Text2.Text = rst.Fields(1)
Text3.Text = rst.Fields(2)
Text4.Text = rst.Fields(3)
Text5.Text = rst.Fields(4)
Text6.Text = rst.Fields(5)
Text7.Text = rst.Fields(6)
End Sub

