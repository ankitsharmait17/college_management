VERSION 5.00
Begin VB.Form Form9 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Faculty"
   ClientHeight    =   9990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14460
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   ScaleHeight     =   9990
   ScaleWidth      =   14460
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Height          =   735
      Left            =   360
      Picture         =   "Form9.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7920
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Create"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6720
      MaskColor       =   &H00004080&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6000
      UseMaskColor    =   -1  'True
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   6480
      TabIndex        =   4
      Top             =   4320
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   6480
      TabIndex        =   3
      Top             =   2880
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   2100
      Left            =   2160
      Picture         =   "Form9.frx":059A
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2460
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Faculty Name"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   615
      Left            =   2520
      TabIndex        =   2
      Top             =   4320
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Faculty ID"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   615
      Left            =   2640
      TabIndex        =   1
      Top             =   3000
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Faculty Account Details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   1095
      Left            =   4920
      TabIndex        =   0
      Top             =   240
      Width           =   7095
   End
   Begin VB.Image Image2 
      Height          =   9405
      Left            =   0
      Picture         =   "Form9.frx":1622
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15900
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim cmd As New ADODB.Command
Dim rst As New ADODB.Recordset
Dim rst1 As New ADODB.Recordset
Dim updstr As String
Private Sub Command1_Click()
If Text1.Text = "" Or Text2.Text = "" Then
MsgBox ("Enter the data! ")
Exit Sub
End If
cmd.CommandText = "select * from faculty where Fac_id='" & Text1.Text & "'"
rst.Close
rst.Open cmd

Do While Not rst.EOF
c = c + 1

rst.MoveNext

Loop


If c <> 0 Then
MsgBox "User id already exists", vbExclamation, "Duplicate"
Text1.Text = " "
Text1.SetFocus
rst.Close
Else
With cmd
    .ActiveConnection = conn
    .CommandText = "Insert into Faculty values('" & Text1.Text & "', '" & Text2.Text & "', '', '', '','','' )"
    .CommandType = adCmdText
    conn.BeginTrans
    .Execute
    conn.CommitTrans
End With
 MsgBox ("Details successfully Created!")
 'conn.Close
 Text1.Text = ""
 Text2.Text = ""
 'conn.Open
 'rst.Close
End If


End Sub

Private Sub Command2_Click()
'rst.Close
'conn.Close
Unload Me
Unload Form2
Load Form2
Form2.Show

End Sub

Private Sub Form_Load()
conn.ConnectionString = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\" & "vb1.mdb; Mode= Read|Write"
conn.CursorLocation = adUseClient
conn.Open
With cmd
.ActiveConnection = conn
.CommandText = "Select * from Faculty"
.CommandType = adCmdText

End With

With rst
.CursorType = adOpenStatic
.CursorLocation = adUseClient
.LockType = adLockOptimistic
.Open cmd
End With
With rst1
.CursorType = adOpenStatic
.CursorLocation = adUseClient
.LockType = adLockOptimistic
End With

End Sub

