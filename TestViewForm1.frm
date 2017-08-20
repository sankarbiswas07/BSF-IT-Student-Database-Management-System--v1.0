VERSION 5.00
Begin VB.Form TestViewForm1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Student Test Entry"
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   10305
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
      ItemData        =   "TestViewForm1.frx":0000
      Left            =   1920
      List            =   "TestViewForm1.frx":0010
      TabIndex        =   22
      Text            =   "I"
      Top             =   1680
      Width           =   1095
   End
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
      ItemData        =   "TestViewForm1.frx":0024
      Left            =   1920
      List            =   "TestViewForm1.frx":0034
      TabIndex        =   3
      Text            =   "I"
      Top             =   1200
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
      ItemData        =   "TestViewForm1.frx":0048
      Left            =   1920
      List            =   "TestViewForm1.frx":004A
      TabIndex        =   2
      Top             =   720
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
      Left            =   1920
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton ComL 
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
      Left            =   8520
      TabIndex        =   18
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton ComP 
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
      Left            =   8520
      TabIndex        =   17
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton ComN 
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
      Left            =   8520
      TabIndex        =   16
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton ComF 
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
      Left            =   8520
      TabIndex        =   15
      Top             =   720
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
      Left            =   1920
      TabIndex        =   6
      Top             =   3480
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
      Left            =   1920
      MaxLength       =   50
      TabIndex        =   4
      Top             =   2400
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
      Left            =   1920
      MaxLength       =   50
      TabIndex        =   5
      Top             =   2880
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   975
      Left            =   120
      TabIndex        =   11
      Top             =   4920
      Width           =   10095
      Begin VB.CommandButton Button 
         BackColor       =   &H00C0C0C0&
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
         Height          =   375
         Index           =   0
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   380
         Width           =   975
      End
      Begin VB.CommandButton Button 
         BackColor       =   &H00C0C0C0&
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
         Height          =   375
         Index           =   1
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   380
         Width           =   975
      End
      Begin VB.CommandButton Button 
         BackColor       =   &H00C0C0C0&
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
         Height          =   375
         Index           =   2
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   380
         Width           =   975
      End
      Begin VB.CommandButton Button 
         BackColor       =   &H00C0C0C0&
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
         Height          =   375
         Index           =   3
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Button 
         BackColor       =   &H00C0C0C0&
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
         Height          =   375
         Index           =   5
         Left            =   9000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
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
      Height          =   255
      Left            =   360
      TabIndex        =   23
      Top             =   1800
      Width           =   1575
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
      Height          =   255
      Left            =   360
      TabIndex        =   21
      Top             =   1320
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
      Height          =   255
      Left            =   360
      TabIndex        =   20
      Top             =   3600
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
      Height          =   255
      Left            =   360
      TabIndex        =   19
      Top             =   360
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
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   840
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
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   2520
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
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   3000
      Width           =   1575
   End
End
Attribute VB_Name = "TestViewForm1"
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
    RS.Open "select max(rno) from TestMarks", Conn
    RNo = IIf(IsNull(RS(0)), 0, RS(0)) + 1
    'MsgBox "insert into TestMarks values('" & Val(Text1) & "' , '" & UCase(Text2) & "', '" & (Text3) & "','" & (Text4) & "','" & (Text5) & "','" & (Text6) & "','" & (Text7) & "', #" & DateFormat(bDate) & "#,'" & Combo1 & "','" & Combo2 & "','" & Combo3 & "')"
    Conn.Execute "insert into TestMarks values(" & RNo & ",'" & Combo1 & "','" & Combo2 & "','" & Combo3 & "','" & Combo4 & "','" & Text1 & "' ,  " & Val(Text3) & ")"
    Button(0).Enabled = True
    Button(1).Enabled = False
Case 2
    If Text1 = "" Then
    MsgBox "please enter Subject Name"
    Text1.SetFocus
    End If

    Conn.Execute "update TestMarks set dName = '" & Combo1 & "',subName = '" & Combo2 & "',sem = '" & Combo3 & "',testNo = '" & Combo4 & "',totAtt=" & Val(Text3) & " where rNo=" & RNo & ""
    
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
    Conn.Execute "delete from TestMarks where rNo=" & RNo & ""
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
RSF.Open "select * from TestMarks order by RegNo", Conn, adOpenDynamic

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
AdminForm.Show
Unload Me

End Sub



Private Sub Form_Load()
Me.Left = 750
Me.Top = 375
Me.Height = 6420
Me.Width = 10425

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

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text2.SetFocus
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
