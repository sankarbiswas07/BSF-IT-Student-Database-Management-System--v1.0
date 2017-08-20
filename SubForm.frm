VERSION 5.00
Begin VB.Form SubForm 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Subject Details"
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "SubForm.frx":0000
   ScaleHeight     =   6015
   ScaleWidth      =   10305
   Begin VB.TextBox Text7 
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
      Left            =   6120
      TabIndex        =   7
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox Text6 
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
      Left            =   6120
      TabIndex        =   6
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox Text5 
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
      Left            =   2160
      TabIndex        =   5
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox Text4 
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
      Left            =   2160
      TabIndex        =   4
      Top             =   2760
      Width           =   975
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
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2520
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
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1920
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
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1320
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
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   16
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
      Left            =   2160
      TabIndex        =   3
      Top             =   1920
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
      Left            =   2160
      MaxLength       =   50
      TabIndex        =   1
      Top             =   720
      Width           =   3495
   End
   Begin VB.TextBox Text2 
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
      Left            =   2160
      MaxLength       =   50
      TabIndex        =   2
      Top             =   1320
      Width           =   3495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00004000&
      Height          =   975
      Left            =   120
      TabIndex        =   12
      Top             =   4560
      Width           =   10095
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
         Height          =   375
         Index           =   0
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   380
         Width           =   975
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
         Height          =   375
         Index           =   1
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   380
         Width           =   975
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
         Height          =   375
         Index           =   2
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   380
         Width           =   975
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
         Height          =   375
         Index           =   3
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   360
         Width           =   975
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
         Height          =   375
         Index           =   5
         Left            =   9000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Min Main Marks"
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
      Left            =   3960
      TabIndex        =   23
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Main Exam Marks"
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
      Left            =   3960
      TabIndex        =   22
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Min Internal Marks"
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
      TabIndex        =   21
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Int Marks"
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
      Left            =   360
      TabIndex        =   20
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total No Classes"
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
      TabIndex        =   15
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject  Name"
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
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Instructor Name"
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
      Top             =   1440
      Width           =   1575
   End
End
Attribute VB_Name = "SubForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Button_Click(Index As Integer)
Select Case Index
Case 0
    Text1 = ""
    Text2 = ""
    Text1.SetFocus
    Button(0).Enabled = False
    Button(1).Enabled = True
Case 1
    If Text1 = "" Then
    MsgBox "please enter Dept Name"
    Text1.SetFocus
    End If
    Conn.Execute "insert into SubMas(subName,IntName,tClass,intMax,intMin,MainMax,MainMin) values('" & UCase(Text1) & "' , '" & (Text2) & "', '" & (Text3) & "', " & Val(Text4) & ", " & Val(Text5) & ", " & Val(Text6) & ", " & Val(Text7) & ")"
    Button(0).Enabled = True
    Button(1).Enabled = False
Case 2
    If Text1 = "" Then
    MsgBox "please enter Subject Name"
    Text1.SetFocus
    End If

    Conn.Execute "update SubMas set intName = '" & Text2 & "',tClass = '" & Text3 & "',intMax=" & Val(Text4) & ",intMin=" & Val(Text5) & ",mainMax=" & Val(Text6) & ",MainMin=" & Val(Text7) & " where subName='" & Text1 & "'"
    
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
    Conn.Execute "delete from SubMas where subName='" & Text1.Text & "'"
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
Text1 = RSF(0)
Text2 = RSF(1)
Text3 = RSF(2)
Text4 = RSF(3)
Text5 = RSF(4)
Text6 = RSF(5)
Text7 = RSF(6)
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
RSF.Open "select * from SubMas order by subName", Conn, adOpenDynamic

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text2.SetFocus
End If

End Sub

Private Sub Text1_LostFocus()
Text1 = UCase(Text1)
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Button(1).SetFocus
End If
End Sub

Private Sub Text2_LostFocus()
Text2 = CheckChar(Text2)
End Sub
