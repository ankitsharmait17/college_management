VERSION 5.00
Begin VB.Form Form10 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Form10"
   ClientHeight    =   9990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14460
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   ScaleHeight     =   9990
   ScaleWidth      =   14460
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Height          =   735
      Left            =   240
      Picture         =   "Form10.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   120
      Width           =   735
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
      Left            =   7560
      TabIndex        =   9
      Top             =   1440
      Width           =   4695
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFF80&
      Height          =   480
      Left            =   7560
      TabIndex        =   8
      Top             =   5655
      Width           =   4695
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFF80&
      Height          =   480
      Left            =   7560
      TabIndex        =   7
      Top             =   4935
      Width           =   4695
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFF80&
      Height          =   480
      Left            =   7560
      TabIndex        =   6
      Top             =   4215
      Width           =   4695
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFF80&
      Height          =   480
      Left            =   7560
      TabIndex        =   5
      Top             =   3495
      Width           =   4695
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00FFFF80&
      Height          =   480
      Left            =   7560
      TabIndex        =   4
      Top             =   2775
      Width           =   4695
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00FFFF80&
      Height          =   480
      Left            =   7560
      TabIndex        =   3
      Top             =   2055
      Width           =   4695
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00FFFF80&
      Height          =   480
      Left            =   7560
      TabIndex        =   2
      Top             =   7095
      Width           =   4695
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00FFFF80&
      Height          =   480
      Left            =   7560
      TabIndex        =   1
      Top             =   6375
      Width           =   4695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Create"
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
      Left            =   6720
      MaskColor       =   &H000080FF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8160
      Width           =   1935
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Registration"
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
      TabIndex        =   19
      Top             =   240
      Width           =   5415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Name"
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
      Top             =   1545
      Width           =   4695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Roll  number"
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
      TabIndex        =   17
      Top             =   2160
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
      Height          =   375
      Left            =   2280
      TabIndex        =   16
      Top             =   2880
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
      Left            =   2280
      TabIndex        =   15
      Top             =   3600
      Width           =   4695
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "GUARDIAN RELATION"
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
      TabIndex        =   14
      Top             =   5040
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
      Left            =   2280
      TabIndex        =   13
      Top             =   5760
      Width           =   4695
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFC0&
      Caption         =   "GUARDIAN NAME"
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
      TabIndex        =   12
      Top             =   4320
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
      Left            =   2280
      TabIndex        =   11
      Top             =   6480
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
      Left            =   2280
      TabIndex        =   10
      Top             =   7200
      Width           =   4695
   End
   Begin VB.Image Image1 
      Height          =   9450
      Left            =   0
      Picture         =   "Form10.frx":059A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14370
   End
End
Attribute VB_Name = "Form10"
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
With cmd
    .ActiveConnection = conn
    .CommandText = "Insert into Student values('" & CInt(Text7.Text) & "', '" & Text1.Text & "', '" & Text6.Text & "', '" & Text5.Text & "', '" & Text4.Text & "', '" & Text3.Text & "', '" & Text2.Text & "', '" & CInt(Text9.Text) & "', '" & CInt(Text8.Text) & "' )"
    .CommandType = adCmdText
    conn.BeginTrans
    .Execute
    conn.CommitTrans
End With
 MsgBox ("Details successfully Created!")
 conn.Close
 conn.Open
 
 Text1.Text = ""
 Text2.Text = ""
 Text3.Text = ""
 Text4.Text = ""
 Text5.Text = ""
 Text6.Text = ""
 Text7.Text = ""
 Text8.Text = ""
 Text9.Text = ""
 
End If

    
    
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command2_Click()
conn.Close



Unload Me
Load Form2
Form2.Show

End Sub

Private Sub Form_Load()
conn.ConnectionString = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\" & "vb1.mdb; Mode= Read|Write"
conn.CursorLocation = adUseClient
conn.Open
With cmd
.ActiveConnection = conn
.CommandText = "Select * from Student"
.CommandType = adCmdText

End With

With rst
.CursorType = adOpenStatic
.CursorLocation = adUseClient
.LockType = adLockOptimistic
.Open cmd
End With
End Sub


