VERSION 5.00
Begin VB.Form Form22 
   Caption         =   "Form22"
   ClientHeight    =   9255
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15510
   LinkTopic       =   "Form22"
   ScaleHeight     =   9255
   ScaleWidth      =   15510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Height          =   735
      Left            =   0
      Picture         =   "stu_pass_chnge.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
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
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3960
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFF80&
      Height          =   480
      IMEMode         =   3  'DISABLE
      Left            =   6600
      PasswordChar    =   "*"
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
      Caption         =   "Change Password"
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
      TabIndex        =   6
      Top             =   360
      Width           =   5295
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "New Password"
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
      TabIndex        =   5
      Top             =   2920
      Width           =   4695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Old Password"
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
      TabIndex        =   4
      Top             =   2160
      Width           =   4695
   End
End
Attribute VB_Name = "Form22"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub
Dim conn As New ADODB.Connection
Dim cmd As New ADODB.Command
Dim rst As New ADODB.Recordset
Private Sub Command1_Click()
With cmd
.ActiveConnection = conn
  .CommandText = "SELECT * From Usr where User_id='" & Form1.ui.Text & "'"
.CommandType = adCmdText

  End With
  With rst
.CursorLocation = adUseClient
.CursorType = adOpenStatic
.LockType = adLockOptimistic
.Open cmd
End With
  
  If Text1.Text <> rst.Fields(1) Then
    MsgBox ("Wrong password")
     rst.Close
        Else
    With cmd
        .ActiveConnection = conn
        .CommandText = "update Usr set Passwd='" & Text2.Text & "' where User_id='" & Form1.ui.Text & "'"
        .CommandType = adCmdText
        conn.BeginTrans 'to insert a new row
        .Execute 'to insert the data
        conn.CommitTrans 'to save the data
        End With
        MsgBox ("record updated successfully")
        Text1.Text = ""
        Text2.Text = ""
        rst.Close
    End If
End Sub

Private Sub Command2_Click()
Unload Me
Form3.Show

End Sub

Private Sub Form_Load()
conn.ConnectionString = "Provider=Microsoft.jet.oledb.4.0;Data source =" & App.Path & "\" & "vb1.mdb;Mode=read|write"
conn.CursorLocation = adUseClient
conn.Open

End Sub

