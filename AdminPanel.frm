VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form AdminDashboard 
   Caption         =   "Start Voting"
   ClientHeight    =   4590
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   255
      Left            =   8880
      TabIndex        =   11
      Top             =   360
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
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
   Begin MSAdodcLib.Adodc DBStatus 
      Height          =   330
      Left            =   5280
      Top             =   5280
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\carl\Desktop\VB6\ACS Voting System\ACSVS.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\carl\Desktop\VB6\ACS Voting System\ACSVS.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
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
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      TabIndex        =   10
      Top             =   3960
      Width           =   2415
   End
   Begin VB.CommandButton Refresh 
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   9
      Top             =   3960
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete/Update Candidate"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   8
      Top             =   3960
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete/Update Voter"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   3960
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      TabIndex        =   5
      Top             =   3240
      Width           =   2415
   End
   Begin VB.CommandButton StopVote 
      Caption         =   "Stop Voting"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   4
      Top             =   3240
      Width           =   2415
   End
   Begin VB.CommandButton StartVote 
      Caption         =   "Start Voting"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   3
      Top             =   3240
      Width           =   2415
   End
   Begin VB.CommandButton AddCandidate 
      Caption         =   "Add Candidate"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   3240
      Width           =   2415
   End
   Begin VB.CommandButton AddVoter 
      Caption         =   "Add Voter"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   3240
      Width           =   2415
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   4048
      _Version        =   393216
      AllowUpdate     =   0   'False
      Enabled         =   -1  'True
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
   Begin MSAdodcLib.Adodc stud 
      Height          =   495
      Left            =   5040
      Top             =   5640
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\carl\Desktop\VB6\ACS Voting System\ACSVS.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\carl\Desktop\VB6\ACS Voting System\ACSVS.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
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
   Begin VB.Label Label1 
      Caption         =   "New Users"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   24.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   7695
   End
End
Attribute VB_Name = "AdminDashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub AddCandidate_Click()
AddPosition.Show
End Sub

Private Sub AddVoter_Click()
AddVoterForm.Show
End Sub

Private Sub Command1_Click()
VoteStatus.Show
End Sub

Private Sub Command2_Click()
DUVoterFOrm.Show
End Sub

Private Sub Command3_Click()
DUCandidateForm.Show
End Sub

Private Sub Command4_Click()
Me.Hide
Login.Show
End Sub

Private Sub Form_Load()
MsgBox "Welcome back!"
    DBStatus.RecordSource = "SELECT * FROM tblVotingStatus"
    stud.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\carl\Desktop\VB6\ACS Voting System\ACSVS.mdb;Persist Security Info=False"
    stud.RecordSource = "SELECT TOP 25 voter_id AS [Voter Id], stud_number AS [Student Number],stud_name As [Student Name], stud_course AS [Student Course], stud_year AS [Year Level], date_created AS [Date Created], date_voted AS [Date Voted] FROM tblStudents WHERE is_deleted = '0' ORDER BY date_created"
    Set DataGrid1.DataSource = stud
    Set DataGrid2.DataSource = DBStatus
        If DBStatus.Recordset.Fields("is_started") = "ongoing" Then
          AddCandidate.Visible = False
        Command3.Visible = False
        StartVote.Visible = False
        StopVote.Visible = True
        ElseIf DBStatus.Recordset.Fields("is_started") = "standby" Then
        AddCandidate.Visible = True
        Command3.Visible = True
        StartVote.Visible = True
        StopVote.Visible = False
        End If
stud.Refresh
DataGrid1.Refresh
End Sub

Private Sub Refresh_Click()
stud.Refresh
DataGrid1.Refresh
End Sub

Private Sub StartVote_Click()
DBStatus.Refresh
        With DBStatus.Recordset
            .Fields("is_started") = "ongoing"
            .Update
        End With
        AddCandidate.Visible = False
        Command3.Visible = False
        StartVote.Visible = False
        StopVote.Visible = True
        DBStatus.Refresh
End Sub

Private Sub StopVote_Click()
        DBStatus.Refresh
        With DBStatus.Recordset
            .Fields("is_started") = "standby"
            .Update
        End With
        AddCandidate.Visible = True
        Command3.Visible = True
        StartVote.Visible = True
        StopVote.Visible = False
        DBStatus.Refresh
End Sub
