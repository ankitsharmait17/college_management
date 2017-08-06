VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form12 
   Caption         =   "Form12"
   ClientHeight    =   8745
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13485
   LinkTopic       =   "Form12"
   ScaleHeight     =   8745
   ScaleWidth      =   13485
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Return"
      Height          =   732
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7560
      Width           =   1932
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000C&
      Caption         =   "Update"
      Enabled         =   0   'False
      Height          =   372
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6000
      Width           =   3492
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000C&
      Height          =   372
      Left            =   9720
      TabIndex        =   5
      Top             =   4800
      Visible         =   0   'False
      Width           =   3372
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H8000000C&
      DataMember      =   "dept id,dept name"
      Height          =   315
      Left            =   9480
      TabIndex        =   4
      Text            =   "Select the field"
      Top             =   2280
      Visible         =   0   'False
      Width           =   3372
   End
   Begin VB.TextBox deptid 
      Height          =   732
      Left            =   2400
      TabIndex        =   3
      Text            =   "enter the course code needed to be updated"
      Top             =   7560
      Width           =   3768
   End
   Begin VB.CommandButton prev 
      BackColor       =   &H8000000C&
      Caption         =   "Prev"
      Height          =   870
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4680
      Width           =   1680
   End
   Begin VB.CommandButton nxt 
      BackColor       =   &H8000000C&
      Caption         =   "Next"
      Height          =   840
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H8000000A&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   7560
      MaskColor       =   &H0080C0FF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7560
      Width           =   1572
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "Form12.frx":0000
      Height          =   3015
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   5318
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   13320
      Top             =   7800
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=vb1.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=vb1.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select *  from Course"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Height          =   9405
      Left            =   0
      Picture         =   "Form12.frx":0015
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15900
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "Enter the Member id whose recored is to be modified"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   480
      TabIndex        =   10
      Top             =   5880
      Width           =   3132
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Document Library Management System"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   12000
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim rst As New ADODB.Recordset
Dim cmd As New ADODB.Command

Private Sub deptid_Click()
deptid.Text = ""
End Sub

Private Sub Form_Load()
Combo1.AddItem ("Course_name")
Combo1.AddItem ("Course_obj")
Combo1.AddItem ("Course_desc")

conn.ConnectionString = "Provider=Microsoft.jet.oledb.4.0;Data source =" & App.Path & "\" & "vb1.mdb;Mode=read|write"
conn.CursorLocation = adUseClient
conn.Open
With cmd
.ActiveConnection = conn
.CommandText = "SELECT * from Course"
.CommandType = adCmdText
End With

With rst
.CursorLocation = adUseClient
.CursorType = adOpenStatic
.LockType = adLockOptimistic
.Open cmd
End With


End Sub
Private Sub Combo1_Click()
MsgBox ("enter the new entry")
Text2.Visible = True
Text2.SetFocus
End Sub

Private Sub Command1_Click()

With cmd
.ActiveConnection = conn
If Combo1.Text = "Course_name" Then
.CommandText = "update Course set Course_name = '" & Text2.Text & "' where Course_code='" & deptid.Text & "'"
End If

If Combo1.Text = "Course_obj" Then
.CommandText = "update Course set Course_obj = '" & Text2.Text & "' where Course_code='" & deptid.Text & "'"
End If
If Combo1.Text = "Course_desc" Then
.CommandText = "update Course set Course_desc = '" & Text2.Text & "' where Course_code='" & deptid.Text & "'"
End If
.CommandType = adCmdText
conn.BeginTrans 'to insert a new row
.Execute 'to insert the data
conn.CommitTrans 'to save the data
MsgBox ("record updated successfully")
.CommandText = "select * from Course"
rst.Close
rst.Open cmd
Adodc1.Refresh
DataGrid2.Refresh
   ' rst.AbsolutePosition = rwid
    'Adodc1.Recordset.AbsolutePosition = rwid
End With
deptid.Text = ""
Combo1.Text = ""
Text2.Text = ""

End Sub

Private Sub Command2_Click()
If Len(deptid.Text) = 0 Then
    MsgBox ("Course id is mandatory for update")
    deptid.SetFocus
End If
cmd.CommandText = "select * from course where Course_code='" + deptid.Text + "'"
rst.Close
rst.Open cmd
Do While Not rst.EOF
c = c + 1
rst.MoveNext
Loop
If c <> 0 Then
MsgBox ("Select the feild to be updated")
Combo1.Visible = True
Else
deptid.Text = ""
MsgBox "Incorrect Course Id", vbExclamation, "ERROR"
End If

End Sub

Private Sub Command3_Click()
Unload Me

Form2.Show

End Sub


Private Sub nxt_Click()
On Error GoTo er

If Not rst.EOF Then
    rst.MoveNext
    Adodc1.Recordset.MoveNext
   ' Call showrcd
End If

Exit Sub

er:
MsgBox "Reached end of file,moving to first"
rst.MoveFirst
Adodc1.Recordset.MoveFirst

'Call showrcd
End Sub

Private Sub prev_Click()
On Error GoTo er

If Not rst.EOF Then
    rst.MovePrevious
    Adodc1.Recordset.MovePrevious
    'Call showrcd
End If

Exit Sub

er:
MsgBox ("Reached BOF,moving to last")
rst.MoveLast
Adodc1.Recordset.MoveLast

'Call showrcd

End Sub

Private Sub Text2_Change()
Command1.Enabled = True


End Sub



