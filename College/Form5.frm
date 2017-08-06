VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Form5"
   ClientHeight    =   9990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14460
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   ScaleHeight     =   9990
   ScaleWidth      =   14460
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   495
      Index           =   2
      Left            =   3600
      TabIndex        =   7
      Text            =   "      14th-20th May, 2016 : Sports Week "
      Top             =   5400
      Visible         =   0   'False
      Width           =   11175
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   495
      Index           =   1
      Left            =   3600
      TabIndex        =   6
      Text            =   "      28th April, 2016 : Rotaract Club Cultural Evening"
      Top             =   4800
      Visible         =   0   'False
      Width           =   11175
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0FFC0&
      Caption         =   "admission"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10800
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2880
      Width           =   3735
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FF8080&
      Caption         =   "Log In"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   495
      Index           =   0
      Left            =   3600
      TabIndex        =   3
      Text            =   "      24th April, 2016 : Pravasna Film Festival "
      Top             =   4200
      Visible         =   0   'False
      Width           =   11175
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "placements"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7080
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2880
      Width           =   3735
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "events"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3360
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2880
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "COURSES"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2880
      Width           =   3375
   End
   Begin VB.Image Image6 
      Height          =   8865
      Left            =   0
      Picture         =   "Form5.frx":0000
      Stretch         =   -1  'True
      Top             =   3720
      Visible         =   0   'False
      Width           =   14445
   End
   Begin VB.Image Image5 
      Height          =   8865
      Left            =   0
      Picture         =   "Form5.frx":9DC9
      Stretch         =   -1  'True
      Top             =   3720
      Visible         =   0   'False
      Width           =   14445
   End
   Begin VB.Image Image2 
      Height          =   8865
      Left            =   0
      Picture         =   "Form5.frx":2C3CC
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   14445
   End
   Begin VB.Image Image4 
      Height          =   2055
      Left            =   2640
      Top             =   5280
      Width           =   1935
   End
   Begin VB.Label Label6 
      Height          =   1455
      Left            =   2730
      TabIndex        =   13
      Top             =   6360
      Width           =   9015
   End
   Begin VB.Image Image3 
      Height          =   600
      Left            =   11520
      Picture         =   "Form5.frx":61DCF
      Top             =   960
      Width           =   600
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFC0C0&
      Caption         =   "heritage.edu@gmail.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   495
      Left            =   11520
      TabIndex        =   12
      Top             =   2280
      Width           =   3015
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "+9823459765"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   495
      Left            =   11520
      TabIndex        =   11
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "+033 22570002"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   495
      Left            =   11520
      TabIndex        =   10
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Contact Us"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   495
      Left            =   12240
      TabIndex        =   9
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   $"Form5.frx":62184
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   3975
      Left            =   3480
      TabIndex        =   8
      Top             =   3840
      Visible         =   0   'False
      Width           =   10695
   End
   Begin VB.Image Image1 
      Height          =   2655
      Left            =   0
      Picture         =   "Form5.frx":62361
      Top             =   0
      Width           =   11475
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer
Dim aa As Integer
Dim code As String
Dim conn As New ADODB.Connection
Dim cmd  As New ADODB.Command
Dim rst As New ADODB.Recordset

Private Sub Combo1_Click()
a = Combo2.ListIndex
If a = 0 Then
Combo3.Enabled = True
Combo3.Visible = True
Else
Combo2.Enabled = True
Combo2.Visible = True
End If
End Sub


Private Sub Combo2_Change()
aa = Combo2.ListIndex
End Sub

Private Sub Combo2_Click()
Label6.Caption = Combo2.Text
End Sub

Private Sub Combo3_Click()

Label6.Visible = True



Text1.Text = "blah"
aa = Combo3.ListIndex
If aa = 0 Then
    code = "CSE"
ElseIf aa = 1 Then
    code = "IT"
ElseIf aa = 2 Then
    code = "ECE"
ElseIf aa = 3 Then
    code = "EE"
ElseIf aa = 4 Then
    code = "ME"
ElseIf aa = 5 Then
    code = "CHE"
ElseIf aa = 6 Then
    code = "CE"
End If
Text1.Enabled = True
Text1.Visible = True
Text1.Text = "blah"

conn.ConnectionString = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\" & "vb.mdb; Mode= Read|Write"
conn.CursorLocation = adUseClient
conn.Open
With cmd
.ActiveConnection = conn
.CommandText = "SELECT * from Course"
.CommandType = adCmdText
End With

With rst
.CursorType = adOpenStatic
.CursorLocation = adUseClient
.LockType = adLockOptimistic
.Open cmd
End With
rst.Close
rst.Open "select * from Course where Course_code=code, conn, adOpenStatic, adLockReadOnly"
Text1.Text = rst!Course_name & vbCrLf & "Course Duration: " & rst!Course_obj & "Years" & vbCrLf & rst!Course_desc
conn.Close



End Sub

Private Sub Command1_Click()
Label1.Enabled = False
Text1(0).Enabled = False
Text1(1).Enabled = False
Text1(2).Enabled = False
Text1(0).Visible = False
Text1(1).Visible = False
Text1(2).Visible = False
Image2.Visible = True
Label1.Visible = False
Image6.Visible = True
Image5.Visible = False

End Sub

Private Sub Command2_Click()
Image2.Visible = False
Label1.Enabled = False
Label1.Visible = False
Text1(0).Visible = True
Text1(1).Visible = True
Text1(2).Visible = True
Text1(0).Enabled = True
Text1(1).Enabled = True
Text1(2).Enabled = True
Label6.Visible = False
Image5.Visible = False
Image6.Visible = False
End Sub

Private Sub Command3_Click()

Image2.Visible = False
Text1(0).Enabled = False
Text1(1).Enabled = False
Text1(2).Enabled = False
Text1(0).Visible = False
Text1(1).Visible = False
Text1(2).Visible = False
Image5.Visible = True
Label1.Enabled = False
Label1.Visible = False
Label6.Visible = False
Image6.Visible = False
End Sub

Private Sub Command4_Click()
Unload Me
Load Form1
Form1.Show
End Sub

Private Sub Command5_Click()

Image2.Visible = False
Text1(0).Enabled = False
Text1(1).Enabled = False
Text1(2).Enabled = False
Text1(0).Visible = False
Text1(1).Visible = False
Image5.Visible = False
Image6.Visible = False
Text1(2).Visible = False
Label1.Enabled = True
Label1.Visible = True
Label6.Visible = False
End Sub

Private Sub Form_Load()
'conn.ConnectionString = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\" & "vb.mdb; Mode= Read|Write"
'conn.CursorLocation = adUseClient
'conn.Open
'With cmd
'.ActiveConnection = conn
'.CommandText = "SELECT * from Course"
'.CommandType = adCmdText
'End With

'With rst
'.CursorType = adOpenStatic
'.CursorLocation = adUseClient
'.LockType = adLockOptimistic
'.Open cmd
'End With



End Sub

