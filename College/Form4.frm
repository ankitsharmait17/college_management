VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Form4"
   ClientHeight    =   9990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14460
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   ScaleHeight     =   9990
   ScaleWidth      =   14460
   StartUpPosition =   3  'Windows Default
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
      Left            =   6480
      TabIndex        =   8
      Top             =   2055
      Width           =   4695
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFF80&
      Height          =   480
      Left            =   6480
      TabIndex        =   7
      Top             =   6615
      Width           =   4695
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFF80&
      Height          =   480
      Left            =   6480
      TabIndex        =   6
      Top             =   5855
      Width           =   4695
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFF80&
      Height          =   480
      Left            =   6480
      TabIndex        =   5
      Top             =   5095
      Width           =   4695
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFF80&
      Height          =   480
      Left            =   6480
      TabIndex        =   4
      Top             =   4335
      Width           =   4695
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00FFFF80&
      Height          =   480
      Left            =   6480
      TabIndex        =   3
      Top             =   3575
      Width           =   4695
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00FFFF80&
      Height          =   480
      Left            =   6480
      TabIndex        =   2
      Top             =   2815
      Width           =   4695
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
      TabIndex        =   1
      Top             =   8040
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Height          =   735
      Left            =   0
      Picture         =   "Form4.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   735
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
      Left            =   1440
      TabIndex        =   16
      Top             =   2160
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
      Left            =   1440
      TabIndex        =   15
      Top             =   2920
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
      Left            =   1440
      TabIndex        =   14
      Top             =   3680
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
      Left            =   1440
      TabIndex        =   13
      Top             =   4440
      Width           =   4695
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "guardian relation"
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
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFC0&
      Caption         =   "course"
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
      Top             =   6720
      Width           =   4695
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Guardian name"
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
      Top             =   5200
      Width           =   4695
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Update  Profile"
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
      TabIndex        =   9
      Top             =   360
      Width           =   5295
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim cmd  As New ADODB.Command
Dim rst As New ADODB.Recordset
Dim updstr As String

Private Sub Command1_Click()
Dim a, b, c, d, e, f, g, h, i As String
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7.Text = "" Or Text8.Text = "" Or Text9.Text = "" Then
    MsgBox ("Information missing! Please enter all the fields.")
    Exit Sub
Else
    a = Text1.Text
    b = Text2.Text
    c = Text3.Text
    d = Text4.Text
    e = Text5.Text
    f = Text6.Text
    g = Text7.Text
    h = Text8.Text
    i = Text9.Text
    updstr = "update Student set Name= b, Addr=c, Cont=d, G_name=e, G_relation=f, Course=g, Year=h, Sem=i where Roll=a"
   
    cmd.CommandText = updstr
    MsgBox ("Details successfully Updated!")
    



End If

    
    
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command2_Click()
'rst.Close
 '           conn.Close
 Unload Form5
            Unload Me
            Form5.Show
End Sub

Private Sub Form_Load()
conn.ConnectionString = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\" & "vb.mdb; Mode= Read|Write"
conn.CursorLocation = adUseClient
conn.Open
With cmd
.ActiveConnection = conn
.CommandType = adCmdText
.Execute
End With

With rst
.CursorType = adOpenStatic
.CursorLocation = adUseClient
.LockType = adLockOptimistic
.Open cmd
End With
End Sub

Private Sub Label10_Click()
End Sub

