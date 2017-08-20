VERSION 5.00
Begin VB.Form LoginForm 
   Caption         =   "Login Form"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9600
   Icon            =   "LoginForm.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "LoginForm.frx":474C7
   ScaleHeight     =   5070
   ScaleWidth      =   9600
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2520
      Picture         =   "LoginForm.frx":E6389
      ScaleHeight     =   465
      ScaleWidth      =   825
      TabIndex        =   6
      Top             =   3120
      Width           =   855
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3480
      Picture         =   "LoginForm.frx":E8477
      ScaleHeight     =   465
      ScaleWidth      =   825
      TabIndex        =   5
      Top             =   3120
      Width           =   855
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6000
      Picture         =   "LoginForm.frx":EA1D1
      ScaleHeight     =   465
      ScaleWidth      =   825
      TabIndex        =   4
      Top             =   3120
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6960
      Picture         =   "LoginForm.frx":EC2BF
      ScaleHeight     =   465
      ScaleWidth      =   825
      TabIndex        =   3
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2400
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2640
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   0
      Top             =   2640
      Width           =   2655
   End
End
Attribute VB_Name = "LoginForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Picture1_Click()
If RS.State = 1 Then RS.Close
RS.Open "select regNo from stuMas where regno='" & Text1 & "'", Conn
If RS.EOF = False Then
RegNoVar = Text1
StuMainForm.Show
Unload Me
Else
MsgBox ("RegNo  is not correct please check")
Text1.SetFocus
End If
End Sub

Private Sub Picture2_Click()
End
End Sub

Private Sub Picture3_Click()
If RS.State = 1 Then RS.Close
RS.Open "select * from logintab where uname='" & Text2 & "' and pword='" & Text3 & "'", Conn
If RS.EOF = False Then
MDIForm1.Show
Unload Me
Else
MsgBox ("username  is not correct please check")
Text1.SetFocus
End If
End Sub

Private Sub Picture4_Click()
End
End Sub
