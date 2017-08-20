VERSION 5.00
Begin VB.Form TestViewForm 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Student Test Entry"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "TestViewForm.frx":0000
   ScaleHeight     =   5625
   ScaleWidth      =   10410
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   10200
      TabIndex        =   21
      Text            =   "Combo3"
      Top             =   7320
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   10320
      TabIndex        =   20
      Text            =   "Combo1"
      Top             =   6840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ComboBox Combo4 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "TestViewForm.frx":415E5
      Left            =   1800
      List            =   "TestViewForm.frx":415F5
      TabIndex        =   18
      Text            =   "I"
      Top             =   240
      Width           =   1095
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "TestViewForm.frx":41609
      Left            =   1800
      List            =   "TestViewForm.frx":4160B
      TabIndex        =   1
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton ComL 
      BackColor       =   &H00FF8080&
      Caption         =   "Last"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3240
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton ComP 
      BackColor       =   &H00FF8080&
      Caption         =   "Previous"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2400
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton ComN 
      BackColor       =   &H00FF8080&
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1560
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton ComF 
      BackColor       =   &H00FF8080&
      Caption         =   "First"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   600
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      MaxLength       =   50
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      MaxLength       =   50
      TabIndex        =   3
      Top             =   2040
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00004000&
      Height          =   1215
      Left            =   240
      TabIndex        =   9
      Top             =   4080
      Width           =   9735
      Begin VB.CommandButton Button 
         BackColor       =   &H00FF8080&
         Caption         =   "&New"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton Button 
         BackColor       =   &H00FF8080&
         Caption         =   "&Save"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   1
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton Button 
         BackColor       =   &H00FF8080&
         Caption         =   "&Modify"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton Button 
         BackColor       =   &H00FF8080&
         Caption         =   "&Delete"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton Button 
         BackColor       =   &H00FF8080&
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   5
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "TestViewForm.frx":4160D
      Top             =   3240
      Width           =   4470
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Test No"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Scored Marks"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Reg No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2160
      Width           =   1575
   End
End
Attribute VB_Name = "TestViewForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Button_Click(Index As Integer)
Select Case Index



Case 5
    Unload Me
    
End Select

If RSF.State = 1 Then RSF.Close
RSF.Open "select * from TestMarks order by RegNo", Conn, adOpenDynamic

End Sub


Private Sub Combo2_Change()
DisRecord
End Sub

Private Sub Combo4_LostFocus()
DisRecord
End Sub

Private Sub ComF_Click()
If RSF.EOF = False Then
RSF.MoveFirst
DisRecord
End If
End Sub
Sub DisRecord()
On Error Resume Next
If RSF.State = 1 Then RSF.Close
RSF.Open "select * from TestMarks where RegNo='" & RegNoVar & "' and testNo='" & Combo4 & "' and subName='" & Combo2 & "'", Conn
If RSF.EOF = False Then
RNo = RSF(0)
Combo1 = RSF(1)
Combo2 = RSF(2)
Combo3 = RSF(3)
Text1 = RSF(5)
Text3 = RSF(6)
    If RS.State = 1 Then RS.Close
    RS.Open "select sName from StuMas where regno='" & RegNoVar & "'", Conn
    If RS.EOF = False Then
    Text2 = RS(0)
    End If
Else
Text2 = ""
Text3 = ""
MsgBox "Test results are not available"
End If
End Sub

Private Sub ComL_Click()
If RSF.EOF = False Then
RSF.MoveLast
DisRecord
End If
End Sub

Private Sub ComN_Click()
If RSF.EOF = False Then
RSF.MoveNext
DisRecord
End If
End Sub

Private Sub ComP_Click()
If RSF.BOF = False Then
RSF.MovePrevious
DisRecord
End If
End Sub

Private Sub Form_Deactivate()
MDIForm1.Show
Unload Me

End Sub



Private Sub Form_Load()
Me.Left = 750
Me.Top = 375
Me.Height = 6420
Me.Width = 10425

Text1 = RegNoVar

If RSF.State = 1 Then RSF.Close
RSF.Open "select * from TestMarks order by RegNo", Conn, adOpenDynamic

If RS.State = 1 Then RS.Close
RS.Open "select dName from DepMas", Conn
Combo1.Clear
Do While Not RS.EOF
Combo1.AddItem RS(0)
RS.MoveNext
Loop
If RS.State = 1 Then RS.Close
RS.Open "select subName from subMas", Conn
Combo2.Clear
Do While Not RS.EOF
Combo2.AddItem RS(0)
RS.MoveNext
Loop
End Sub

Private Sub Image1_Click()
DisRecord
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
DisRecord
End If

End Sub

Private Sub Text1_LostFocus()
If RS.State = 1 Then RS.Close
RS.Open "select sName from StuMas where regno='" & Text1 & "'", Conn
If RS.EOF = True Then
MsgBox "This Register No does not exit Please check"
Else
Text2 = RS(0)
DisRecord
End If

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Button(1).SetFocus
End If
End Sub

Private Sub Text2_LostFocus()
Text2 = CheckChar(Text2)
End Sub
