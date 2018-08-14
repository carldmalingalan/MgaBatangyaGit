VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form AddPosition 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   7425
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc DBPosition 
      Height          =   330
      Left            =   240
      Top             =   720
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
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   24.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3960
      TabIndex        =   6
      Top             =   3480
      Width           =   2175
   End
   Begin VB.CommandButton AddNewPos 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   24.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      TabIndex        =   5
      Top             =   3480
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   2520
      Width           =   3135
   End
   Begin VB.ComboBox Position 
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      ItemData        =   "AddPosition.frx":0000
      Left            =   3000
      List            =   "AddPosition.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1440
      Width           =   3135
   End
   Begin VB.Label Label3 
      Caption         =   "Name :"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Position :"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Alter Position"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   6735
   End
End
Attribute VB_Name = "AddPosition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AddNewPos_Click()
Dim table As String
    If Not Position.Text = "" And Not Text1.Text = "" Then
        table = "tbl" & Position.Text
        DBPosition.RecordSource = "SELECT * FROM " & table
        DBPosition.Refresh
        With DBPosition.Recordset
            .AddNew
            .Fields("candidate_name") = Text1.Text
            .Update
        End With
    Else
    MsgBox "Please fill all fields.", , "Invalid Parameters."
    End If
    Text1.Text = ""
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Me.Hide
End Sub

Private Sub Form_Load()
DBPosition.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\carl\Desktop\VB6\ACS Voting System\ACSVS.mdb;Persist Security Info=False"
Text1.Text = ""
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    Dim strPat As RegExp
    Set strPat = New RegExp
    strPat.Pattern = "[^a-zA-z.]"
    If strPat.Test(Chr$(KeyAscii)) And Not KeyAscii = 8 And Not KeyAscii = 32 Then
        KeyAscii = 0
    End If
End Sub
