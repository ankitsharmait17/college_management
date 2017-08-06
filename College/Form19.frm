VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form19 
   Caption         =   "Form19"
   ClientHeight    =   8280
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13860
   LinkTopic       =   "Form19"
   ScaleHeight     =   8280
   ScaleWidth      =   13860
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   492
      Left            =   2160
      TabIndex        =   5
      Text            =   "Subject Code"
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Return"
      Height          =   732
      Left            =   7800
      TabIndex        =   3
      Top             =   7200
      Width           =   2532
   End
   Begin VB.TextBox Text1 
      Height          =   492
      Left            =   480
      TabIndex        =   2
      Text            =   "Enter the roll"
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      Height          =   732
      Left            =   4560
      TabIndex        =   1
      Top             =   7200
      Width           =   2532
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Marks Entry"
      Height          =   732
      Left            =   10920
      TabIndex        =   0
      Top             =   7200
      Width           =   2532
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   1080
      Top             =   960
      Visible         =   0   'False
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      RecordSource    =   "SELECT *  from marks"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form19.frx":0000
      Height          =   4335
      Left            =   600
      TabIndex        =   4
      Top             =   2160
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   7646
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
   Begin VB.Image Image1 
      Height          =   9525
      Left            =   0
      Picture         =   "Form19.frx":0015
      Stretch         =   -1  'True
      Top             =   360
      Width           =   23580
   End
End
Attribute VB_Name = "Form19"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim cmd As New ADODB.Command
Dim rst As New ADODB.Recordset

Private Sub Command1_Click()
If rst.State <> adStateClosed Then rst.Close
If conn.State <> adStateClosed Then conn.Close
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
   App.Path & "\" & "vb1.mdb;Mode=Read|Write"
    conn.CursorLocation = adUseClient
    conn.Open

With cmd
.ActiveConnection = conn
  .CommandText = "SELECT * From Marks where roll='" & Text1.Text & "' and Subject_code='" & Text2.Text & "'"
.CommandType = adCmdText
  End With

With rst
    .CursorType = adOpenStatic
    .CursorLocation = adUseClient
    .LockType = adLockOptimistic
    .Open cmd
    
    
End With

Do While Not rst.EOF
c = c + 1
rst.MoveNext
Loop
If c = 0 Then
MsgBox ("No marks entered for this roll in this subject ")
Else
Adodc1.Refresh
DataGrid1.Refresh
Adodc1.RecordSource = "SELECT * From Marks where roll='" & Text1.Text & "' and Subject_code='" & Text2.Text & "'"
conn.Close
End If
End Sub

Private Sub Command3_Click()
If rst.State <> adStateClosed Then rst.Close
If conn.State <> adStateClosed Then conn.Close

Unload Me
Load Form3
Form3.Show

End Sub

Private Sub Command4_Click()
Me.Hide
Form20.Show
End Sub

Private Sub Form_Load()
conn.ConnectionString = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\" & "vb1.mdb; Mode= Read|Write"
conn.CursorLocation = adUseClient
conn.Open
With cmd
.ActiveConnection = conn
.CommandText = "SELECT * from Marks"
.CommandType = adCmdText
End With
With rst

.CursorType = adOpenStatic
.CursorLocation = adUseClient
.LockType = adLockOptimistic
.Open cmd
End With

End Sub

Private Sub Text1_Click()
Text1.Text = ""
End Sub




