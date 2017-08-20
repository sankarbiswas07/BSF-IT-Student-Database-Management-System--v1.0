VERSION 5.00
Begin VB.Form AdminForm 
   Caption         =   "Admin Form"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10305
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   10305
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label8 
      Caption         =   "Report"
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
      Left            =   600
      TabIndex        =   7
      Top             =   4560
      Width           =   5055
   End
   Begin VB.Label Label7 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   8520
      TabIndex        =   6
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Main Exam Marks Entry"
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
      Left            =   600
      TabIndex        =   5
      Top             =   3960
      Width           =   5055
   End
   Begin VB.Label Label5 
      Caption         =   "Int Exam Marks Entry"
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
      Left            =   600
      TabIndex        =   4
      Top             =   3360
      Width           =   5055
   End
   Begin VB.Label Label4 
      Caption         =   "Attendance Entry"
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
      Left            =   600
      TabIndex        =   3
      Top             =   2760
      Width           =   5055
   End
   Begin VB.Label Label2 
      Caption         =   "Subject Entry"
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
      Left            =   600
      TabIndex        =   2
      Top             =   1560
      Width           =   5055
   End
   Begin VB.Label Label1 
      Caption         =   "Department Entry"
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
      Left            =   600
      TabIndex        =   1
      Top             =   960
      Width           =   5055
   End
   Begin VB.Label Label3 
      Caption         =   "Student Registration"
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
      Left            =   600
      TabIndex        =   0
      Top             =   2160
      Width           =   5055
   End
End
Attribute VB_Name = "AdminForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'Conn.Open "Provider=MSDAORA.1;User ID=scott;password=tiger;Persist Security Info=true"   'oracle
'Conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\sData.mdb;Persist Security Info=False" 'msaccess
Me.Left = 750
Me.Top = 375
Me.Height = 6420
Me.Width = 10425
End Sub

Private Sub Label1_Click()
DepForm.Show
End Sub

Private Sub Label2_Click()
SubForm.Show
End Sub

Private Sub Label3_Click()
StuEntryForm.Show
End Sub

Private Sub Label4_Click()
AttEntryForm.Show
End Sub

Private Sub Label5_Click()
TestEForm.Show
End Sub

Private Sub Label6_Click()
ExamEForm.Show
End Sub

Private Sub Label7_Click()
End
End Sub

Private Sub Label8_Click()
ReportForm.Show
End Sub
