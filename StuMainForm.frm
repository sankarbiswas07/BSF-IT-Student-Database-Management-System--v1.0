VERSION 5.00
Begin VB.Form StuMainForm 
   BackColor       =   &H00FF80FF&
   Caption         =   "Student Main Scre"
   ClientHeight    =   10950
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   Picture         =   "StuMainForm.frx":0000
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer3 
      Interval        =   60000
      Left            =   1000
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Interval        =   70
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   1680
      Picture         =   "StuMainForm.frx":4F3842
      ScaleHeight     =   585
      ScaleWidth      =   3345
      TabIndex        =   8
      Top             =   9120
      Width           =   3375
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   1560
      Picture         =   "StuMainForm.frx":4FAAB4
      ScaleHeight     =   585
      ScaleWidth      =   3465
      TabIndex        =   7
      Top             =   7680
      Width           =   3495
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   15720
      Picture         =   "StuMainForm.frx":501856
      ScaleHeight     =   465
      ScaleWidth      =   4425
      TabIndex        =   6
      Top             =   5640
      Width           =   4455
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   15720
      Picture         =   "StuMainForm.frx":508898
      ScaleHeight     =   465
      ScaleWidth      =   4425
      TabIndex        =   5
      Top             =   4680
      Width           =   4455
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   15720
      Picture         =   "StuMainForm.frx":50F8DA
      ScaleHeight     =   465
      ScaleWidth      =   4425
      TabIndex        =   4
      Top             =   3720
      Width           =   4455
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   15720
      Picture         =   "StuMainForm.frx":51691C
      ScaleHeight     =   465
      ScaleWidth      =   4425
      TabIndex        =   3
      Top             =   2760
      Width           =   4455
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   13440
      Top             =   4560
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "B.S.F. IT STUDENT DATABASE MANAGEMENT SYSTEM v1.0"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   555
      Left            =   120
      TabIndex        =   9
      Top             =   10320
      Width           =   20160
   End
   Begin VB.Label Label1c 
      BackColor       =   &H00FF80FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   495
      Left            =   15840
      TabIndex        =   2
      Top             =   9360
      Width           =   4095
   End
   Begin VB.Label Label1b 
      BackColor       =   &H00FF80FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   495
      Left            =   15840
      TabIndex        =   1
      Top             =   8400
      Width           =   4095
   End
   Begin VB.Label Label1a 
      BackColor       =   &H00FF80FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   495
      Left            =   15840
      TabIndex        =   0
      Top             =   7440
      Width           =   4095
   End
End
Attribute VB_Name = "StuMainForm"
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
Label1a.Caption = "Reg No : " & RegNoVar
Label1b.Caption = "Date : " & Format(Date, "dd/MMM/yyyy")

End Sub



Private Sub OLE1_Updated(Code As Integer)

End Sub

Private Sub Picture1_Click()
StuViewForm.Show
End Sub

Private Sub Picture2_Click()
AttViewForm.Show
End Sub

Private Sub Picture3_Click()
TestViewForm.Show
End Sub

Private Sub Picture4_Click()
ResultForm.Show
End Sub

Private Sub Picture5_Click()
LoginForm.Show
Unload Me
End Sub

Private Sub Picture6_Click()
End
End Sub

Private Sub Timer1_Timer()
Label1c = "Time : " & Format(Now, "HH:MM:SS")
End Sub

Private Sub Timer2_Timer()
If Label1.Left < 23000 Then
Label1.Left = Label1.Left + 100
Else
Label1.Left = -0
End If
End Sub

Private Sub Timer3_Timer()
Label1.Visible = False
End Sub
