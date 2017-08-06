VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Form1"
   ClientHeight    =   9990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9990
   ScaleWidth      =   14460
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Height          =   735
      Left            =   360
      Picture         =   "vb_10.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   240
      Width           =   735
   End
   Begin VB.Frame f1 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8220
      Left            =   2310
      TabIndex        =   0
      Top             =   885
      Width           =   9840
      Begin VB.TextBox ui 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3720
         TabIndex        =   11
         Top             =   2040
         Width           =   2340
      End
      Begin VB.TextBox pw 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         IMEMode         =   3  'DISABLE
         Left            =   3720
         PasswordChar    =   "*"
         TabIndex        =   10
         Top             =   2760
         Width           =   2340
      End
      Begin VB.OptionButton ad 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Admin"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   0
         Left            =   3720
         TabIndex        =   9
         Top             =   3480
         Width           =   1005
      End
      Begin VB.OptionButton ad 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Faculty"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   1
         Left            =   4920
         TabIndex        =   8
         Top             =   3480
         Width           =   1005
      End
      Begin VB.OptionButton ad 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Student"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Index           =   2
         Left            =   6120
         TabIndex        =   7
         Top             =   3480
         Width           =   1005
      End
      Begin VB.CommandButton submit 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Log In"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   4680
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Log In"
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
         Height          =   855
         Left            =   3840
         TabIndex        =   5
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label gp 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Group"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   585
         Left            =   1200
         TabIndex        =   3
         Top             =   3480
         Width           =   1005
      End
      Begin VB.Label pass 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   405
         Left            =   1200
         TabIndex        =   2
         Top             =   2775
         Width           =   1605
      End
      Begin VB.Label userid 
         BackColor       =   &H00C0FFC0&
         Caption         =   "User Id"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   405
         Left            =   1200
         TabIndex        =   1
         Top             =   2055
         Width           =   1605
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim cmd  As New ADODB.Command
Dim rst As New ADODB.Recordset
Dim n As Integer

Private Sub ad_Click(Index As Integer)
n = Index
End Sub


Private Sub Command2_Click()
rst.Close
conn.Close
Unload Me
Load Form5
Form5.Show

End Sub

Private Sub Form_Load()

conn.ConnectionString = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\" & "vb1.mdb; Mode= Read|Write"
conn.CursorLocation = adUseClient
conn.Open
With cmd
.ActiveConnection = conn
.CommandText = "SELECT * from Usr"
.CommandType = adCmdText
End With

With rst
.CursorType = adOpenStatic
.CursorLocation = adUseClient
.LockType = adLockOptimistic
.Open cmd
End With
End Sub

Private Sub submit_Click()
If ui.Text = "" Then
    MsgBox "Enter User Id!"
    ui.SetFocus
    Exit Sub
ElseIf pw.Text = "" Then
    MsgBox "Enter Password!"
    pw.SetFocus
    Exit Sub
Else
        rst.Close
    '    If rst.State <> adStateClosed Then rst.Close
     '   If conn.State <> adStateClosed Then conn.Close
        rst.Open "select * from Usr where User_id='" & ui.Text & "' and Passwd='" & pw.Text & "'", conn, adOpenStatic, adLockReadOnly
        If rst.RecordCount < 1 Then
            MsgBox ("User Id/password is is invalid ")
            ui.SetFocus
            Exit Sub
        Else
            If rst.Fields(3) = "admin" And ad(0).Value = True Then
            rst.Close
            conn.Close
           ' Unload Me
            Load Form2
            Form2.Show
            Exit Sub
            ElseIf rst.Fields(3) = "Faculty" And ad(1).Value = True Then
           ' rst.Close
            'conn.Close
                      
            'Unload Me
            Load Form3
            Form3.Show
            Exit Sub
            ElseIf rst.Fields(3) = "Student" And ad(2).Value = True Then
            rst.Close
            conn.Close
            Load Form6
            Form6.Show
            Exit Sub
            Else
            MsgBox ("Invalid user type")
            pw.SetFocus
            Exit Sub
            End If
        End If
        Set rst = Nothing
End If
'rst.Close
'conn.Close
End Sub

