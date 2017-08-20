VERSION 5.00
Begin VB.Form ReportForm 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "ReportForm.frx":0000
   ScaleHeight     =   6015
   ScaleWidth      =   10305
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   7680
      Picture         =   "ReportForm.frx":F0C02
      ScaleHeight     =   435
      ScaleWidth      =   2355
      TabIndex        =   5
      Top             =   5280
      Width           =   2415
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Student Marks Card"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   4800
      Width           =   3735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Exam Marks List"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   4080
      Width           =   3855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Test Marks List"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   3360
      Width           =   3855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Attendence detail List"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   2640
      Width           =   3855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Student Detail List"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   1920
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Student Address List"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   1080
      Width           =   3855
   End
End
Attribute VB_Name = "ReportForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Deactivate()
MDIForm1.Show
Unload Me

End Sub



Private Sub Form_Load()
Me.Left = 750
Me.Top = 375
Me.Height = 6420
Me.Width = 10425


End Sub


Private Sub Label1_Click()
If RS.State = 1 Then RS.Close
RS.Open "select * from stuMas order by dName,regno", Conn
Set stuAddRep.DataSource = RS
stuAddRep.Show
End Sub

Private Sub Label2_Click()
If RS.State = 1 Then RS.Close
RS.Open "select * from stuMas order by dName,regno", Conn
Set stuDetRep.DataSource = RS
stuDetRep.Show
End Sub

Private Sub Label3_Click()
If RS.State = 1 Then RS.Close
RS.Open "select * from AttMas order by dName,subName,sem,regno", Conn
Set StuAtteList.DataSource = RS
StuAtteList.Show
End Sub

Private Sub Label4_Click()

TestResultRep.Show
End Sub

Private Sub Label5_Click()
'If RS.State = 1 Then RS.Close
'RS.Open "select * from ExamMarks", Conn
'Set ExamResult.DataSource = RS
ExamResult.Show
End Sub

Private Sub Label6_Click()
Form1.Show
End Sub

Private Sub Picture1_Click()
Unload Me
End Sub
