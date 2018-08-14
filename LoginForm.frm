VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Login 
   Caption         =   "Form1"
   ClientHeight    =   3945
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   10920
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc AuthUser 
      Height          =   330
      Left            =   8520
      Top             =   4320
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
   Begin VB.TextBox StudNum 
      Alignment       =   2  'Center
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
      Left            =   3120
      TabIndex        =   1
      Top             =   3000
      Width           =   4815
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Enter Student Number: "
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
      Left            =   720
      TabIndex        =   2
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Welcome to ACS Voting System"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   39.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   10215
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub StudNum_KeyPress(KeyAscii As Integer)
   Dim strFormat As RegExp
   Set strFormat = New RegExp
   strFormat.Pattern = "[^0-9-]"
   If Len(StudNum.Text) > 10 And KeyAscii = 13 Then
    If StudNum.Text = "2014-102813" Then
        Me.Hide
        AdminLogin.Show
    Else
        AuthUser.RecordSource = "SELECT * FROM tblVotingStatus"
        AuthUser.Refresh
            If AuthUser.Recordset("is_started") = "ongoing" Then
                AuthUser.RecordSource = "SELECT TOP 1 * FROM tblStudents WHERE stud_number = '" & StudNum.Text & "' AND is_deleted = '0' AND is_voted = '0'"
                AuthUser.Refresh
                    If AuthUser.Recordset.EOF Then
                        MsgBox "User not found or voted already!"
                    Else
                        MsgBox "Welcome " & AuthUser.Recordset("stud_name")
                        Dim BF As New BallotForm
                        BF.StudNum = AuthUser.Recordset("stud_number")
                        BF.StudName = AuthUser.Recordset("stud_name")
                        Me.Hide
                        BF.Show
                    End If
            Else
                MsgBox "Sorry you can't vote right now!"
            End If
    End If
    StudNum.Text = ""
   End If
    Select Case KeyAscii
        Case 48 To 57
        Case 45
        Case 8
        Case Else
        KeyAscii = 0
    End Select
    
    If Len(StudNum.Text) > 10 And Not KeyAscii = 8 Then
        KeyAscii = 0
    End If
    If Not Len(StudNum.Text) = 4 And KeyAscii = 45 Then
    KeyAscii = 0
    End If
    If Len(StudNum.Text) = 4 And Not KeyAscii = 45 And Not KeyAscii = 8 Then
    KeyAscii = 0
    End If
    
    
        
End Sub
