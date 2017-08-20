VERSION 5.00
Begin VB.Form AttEntryForm 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Student Attendance Entry"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "AttEntryForm.frx":0000
   ScaleHeight     =   5910
   ScaleWidth      =   10305
   Begin VB.ComboBox Combo3 
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
      ItemData        =   "AttEntryForm.frx":415E5
      Left            =   2040
      List            =   "AttEntryForm.frx":415F5
      TabIndex        =   3
      Text            =   "I"
      Top             =   1320
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
      ItemData        =   "AttEntryForm.frx":41609
      Left            =   2040
      List            =   "AttEntryForm.frx":4160B
      TabIndex        =   2
      Top             =   840
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
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
      Left            =   2040
      TabIndex        =   1
      Top             =   360
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
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2760
      Width           =   1575
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
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2040
      Width           =   1575
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
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1200
      Width           =   1575
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
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   600
      Width           =   1575
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
      Left            =   2040
      TabIndex        =   6
      Top             =   3000
      Width           =   975
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
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   4
      Top             =   1920
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
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   5
      Top             =   2400
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00004000&
      Height          =   1215
      Left            =   120
      TabIndex        =   11
      Top             =   4080
      Width           =   9615
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
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   360
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
         Height          =   495
         Index           =   1
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
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
         Height          =   480
         Index           =   2
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   380
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
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   1095
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
         Height          =   495
         Index           =   5
         Left            =   8280
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Semister"
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
      TabIndex        =   21
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Attendance"
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
      TabIndex        =   20
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Dept Name"
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
      Top             =   480
      Width           =   1575
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
      TabIndex        =   14
      Top             =   960
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
      TabIndex        =   13
      Top             =   2040
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
      TabIndex        =   12
      Top             =   2520
      Width           =   1575
   End
End
Attribute VB_Name = "AttEntryForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Button_Click(Index As Integer)
Select Case Index
Case 0
    Text1 = ""
    Text2 = ""
    Combo1.SetFocus
    Button(0).Enabled = False
    Button(1).Enabled = True
Case 1
    If Text1 = "" Then
    MsgBox "please enter Dept Name"
    Text1.SetFocus
    End If
    
    If RS.State = 1 Then RS.Close
    RS.Open "select max(rno) from attMas", Conn
    RNo = IIf(IsNull(RS(0)), 0, RS(0)) + 1
    'MsgBox "insert into AttMas values('" & Val(Text1) & "' , '" & UCase(Text2) & "', '" & (Text3) & "','" & (Text4) & "','" & (Text5) & "','" & (Text6) & "','" & (Text7) & "', #" & DateFormat(bDate) & "#,'" & Combo1 & "','" & Combo2 & "','" & Combo3 & "')"
    Conn.Execute "insert into AttMas values(" & RNo & ",'" & Combo1 & "','" & Combo2 & "','" & Combo3 & "','" & Text1 & "' ,  " & Val(Text3) & ")"
    Button(0).Enabled = True
    Button(1).Enabled = False
Case 2
    If Text1 = "" Then
    MsgBox "please enter Subject Name"
    Text1.SetFocus
    End If

    Conn.Execute "update AttMas set dName = '" & Combo1 & "',subName = '" & Combo2 & "',sem = '" & Combo3 & "',totAtt=" & Val(Text3) & " where rNo=" & RNo & ""
    
    Button(0).Enabled = True
    Button(1).Enabled = False
    Button(2).Enabled = False
    Button(3).Enabled = False
 
Case 3
    If vbNo = MsgBox(" Do you want delete record ", vbYesNo) Then Exit Sub
    If Text1 = "" Then
    MsgBox ("Please enter the Subject Name")
    Text1.SetFocus
    Exit Sub
    Else
    Conn.Execute "delete from AttMas where rNo=" & RNo & ""
    Text1 = ""
    Text2 = ""
    Button(0).Enabled = True
    Button(1).Enabled = False
    Button(2).Enabled = False
    Button(3).Enabled = False

    End If


Case 5
    Unload Me
    
End Select

If RSF.State = 1 Then RSF.Close
RSF.Open "select * from AttMas order by RegNo", Conn, adOpenDynamic

End Sub

Private Sub Command1_Click()

End Sub

Private Sub ComF_Click()
If RSF.EOF = False Then
RSF.MoveFirst
DisRecord
End If
End Sub
Sub DisRecord()
On Error Resume Next
If RSF.EOF = False Then
RNo = RSF(0)
Combo1 = RSF(1)
Combo2 = RSF(2)
Combo3 = RSF(3)
Text1 = RSF(4)
Text3 = RSF(5)

    Button(2).Enabled = True
    Button(3).Enabled = True
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

If RSF.State = 1 Then RSF.Close
RSF.Open "select * from AttMas order by RegNo", Conn, adOpenDynamic

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

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text3.SetFocus
End If

End Sub

Private Sub Text1_LostFocus()
If RS.State = 1 Then RS.Close
RS.Open "select sName from StuMas where regno='" & Text1 & "'", Conn
If RS.EOF = True Then
MsgBox "This Register No does not exit Please check"
Else
Text2 = RS(0)
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
