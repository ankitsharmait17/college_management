VERSION 5.00
Begin VB.Form Form11 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Form11"
   ClientHeight    =   9990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14460
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   ScaleHeight     =   9990
   ScaleWidth      =   14460
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFF80&
      Height          =   885
      Left            =   7560
      TabIndex        =   17
      Top             =   3360
      Width           =   4575
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFF80&
      Height          =   525
      Left            =   7560
      TabIndex        =   15
      Top             =   2520
      Width           =   4575
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFF80&
      Height          =   525
      Left            =   7560
      TabIndex        =   14
      Top             =   1560
      Width           =   4575
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFF80&
      Height          =   525
      Left            =   7560
      TabIndex        =   13
      Top             =   720
      Width           =   4575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6750
      MaskColor       =   &H000080FF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8708
      Width           =   1935
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00FFFF80&
      Height          =   480
      Left            =   7590
      TabIndex        =   5
      Top             =   7230
      Width           =   4695
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00FFFF80&
      Height          =   480
      Left            =   7590
      TabIndex        =   4
      Top             =   8010
      Width           =   4695
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFF80&
      Height          =   525
      Left            =   7560
      TabIndex        =   3
      Top             =   4680
      Width           =   4575
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00FFFF80&
      Height          =   480
      Left            =   7560
      TabIndex        =   2
      Top             =   6510
      Width           =   4695
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00FFFF80&
      Height          =   480
      Left            =   7560
      TabIndex        =   1
      Top             =   5640
      Width           =   4695
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Height          =   735
      Left            =   120
      Picture         =   "Form11.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   735
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
      Height          =   375
      Left            =   2280
      TabIndex        =   20
      Top             =   3720
      Width           =   4695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "guardian name"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   19
      Top             =   5640
      Width           =   4695
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFC0&
      Caption         =   "user id"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   18
      Top             =   1920
      Width           =   4695
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "name"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   16
      Top             =   2760
      Width           =   4695
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFC0&
      Caption         =   "SEMESTER"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2310
      TabIndex        =   12
      Top             =   7995
      Width           =   4695
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFC0&
      Caption         =   "YEAR"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2310
      TabIndex        =   11
      Top             =   7335
      Width           =   4695
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFC0&
      Caption         =   "COURSE"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2310
      TabIndex        =   10
      Top             =   6555
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
      Height          =   375
      Left            =   2400
      TabIndex        =   9
      Top             =   4800
      Width           =   4695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFC0&
      Caption         =   " Roll  number"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   960
      Width           =   4695
   End
   Begin VB.Label Label10 
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
      Left            =   3840
      TabIndex        =   7
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "Form11"
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
.CommandText = "update Student set [Name]='" & Text3.Text & "',[Addr]='" & Text4.Text & "',[Cont]='" & Text5.Text & "',[G_name]='" & Text6.Text & "',[Course]='" & Text7.Text & "',[Year]='" & Text8.Text & "',[Sem]='" & Text9.Text & "' where [user_id]='" & Text2.Text & "'"
.CommandType = adCmdText
conn.BeginTrans 'to insert a new row
.Execute 'to insert the data
conn.CommitTrans 'to save the data

MsgBox ("record updated successfully")
.CommandText = "select * from Student"
rst.Close
    rst.Open cmd
End With

End Sub

Private Sub Command2_Click()
Unload Me
Load Form6
Form6.Show

End Sub

Private Sub Form_Load()
conn.ConnectionString = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\" & "vb1.mdb; Mode= Read|Write"
conn.CursorLocation = adUseClient
conn.Open
With cmd
.ActiveConnection = conn
.CommandText = "SELECT * from student where user_id='" & Form1.ui.Text & "'"
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
Text8.Text = rst.Fields(7)
Text9.Text = rst.Fields(8)
End Sub


