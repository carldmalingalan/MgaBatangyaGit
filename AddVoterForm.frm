VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form AddVoterForm 
   Caption         =   "Form1"
   ClientHeight    =   2955
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   ScaleHeight     =   2955
   ScaleWidth      =   9090
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Auth 
      Height          =   330
      Left            =   4560
      Top             =   4320
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
   Begin VB.CommandButton Command2 
      Caption         =   "CANCEL"
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
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADD"
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
      TabIndex        =   8
      Top             =   360
      Width           =   2175
   End
   Begin VB.TextBox Text3 
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
      Left            =   2160
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   2280
      Width           =   3015
   End
   Begin MSDataListLib.DataCombo CourseDC 
      Height          =   495
      Left            =   2160
      TabIndex        =   5
      Top             =   1680
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   873
      _Version        =   393216
      Style           =   2
      ListField       =   ""
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Agency FB"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc course 
      Height          =   495
      Left            =   3600
      Top             =   5400
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
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
   Begin VB.TextBox Text2 
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
      Left            =   2160
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   960
      Width           =   3015
   End
   Begin VB.TextBox Text1 
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
      Left            =   2160
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label Label4 
      Caption         =   "Student Year : "
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Student Course : "
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Student Name : "
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Student Number: "
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "AddVoterForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CourseItems As Variant
Private Sub Command1_Click()
Auth.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\carl\Desktop\VB6\ACS Voting System\ACSVS.mdb;Persist Security Info=False"
Auth.RecordSource = "SELECT * FROM tblStudents WHERE stud_number = '" & Text1.Text & "'"
Auth.Refresh
If Auth.Recordset.EOF Then
    If Not Text1.Text = "" And Not Text2.Text = "" And Not Text3.Text = "" And Not CourseDC.Text = "" Then
    Auth.RecordSource = "SELECT * FROM tblStudents"
    Auth.Refresh
        With Auth.Recordset
        .AddNew
        .Fields("stud_number") = Text1.Text
        .Fields("stud_name") = Text2.Text
        .Fields("stud_course") = CourseDC.Text
        .Fields("stud_year") = Text3.Text
        .Fields("date_created") = Now()
        .Update
        End With
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    CourseDC.Text = "Select a course"
    MsgBox "Student created successfully!", vbQuestion, "Success!"
    AdminDashboard.stud.Refresh
    AdminDashboard.DataGrid1.Refresh
    Else
    MsgBox "Please fill the field properly"
    End If
Else
MsgBox "Student is already registered!"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
CourseDC.Text = "Select a course"
End If


End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
CourseDC.Text = "Select a course"
Me.Hide
AdminDashboard.Show
End Sub

Private Sub Form_Load()
course.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\carl\Desktop\VB6\ACS Voting System\ACSVS.mdb;Persist Security Info=False"
course.RecordSource = "SELECT * FROM tblCourses"
Set CourseDC.DataSource = course
Set CourseDC.RowSource = course
CourseDC.ListField = "course_name"

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
CourseDC.Text = "Select a course"
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
        Case 48 To 57
        Case 45
        Case 8
        Case Else
        KeyAscii = 0
    End Select
    
    If Len(Text1.Text) > 10 And Not KeyAscii = 8 Then
        KeyAscii = 0
    End If
    If Not Len(Text1.Text) = 4 And KeyAscii = 45 Then
    KeyAscii = 0
    End If
    If Len(Text1.Text) = 4 And Not KeyAscii = 45 And Not KeyAscii = 8 Then
    KeyAscii = 0
    End If
End Sub


Private Sub Text2_KeyPress(KeyAscii As Integer)
Dim strPat As RegExp
Set strPat = New RegExp
strPat.Pattern = "[^a-zA-z.]"
    If strPat.Test(Chr$(KeyAscii)) And Not KeyAscii = 8 And Not KeyAscii = 32 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 48 To 57
    Case 8
    Case Else
    KeyAscii = 0
    End Select
    
    If Len(Text3.Text) >= 2 And Not KeyAscii = 8 Then
    KeyAscii = 0
    End If
End Sub
