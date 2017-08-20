VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form ResultForm 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Student Test Entry"
   ClientHeight    =   9600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9600
   ScaleWidth      =   13800
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   8295
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   14631
      _Version        =   393216
      Rows            =   10
      Cols            =   5
      FixedRows       =   0
      FixedCols       =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   8520
      Width           =   13575
      Begin VB.CommandButton Command1 
         Caption         =   "Print"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1455
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
         Left            =   11760
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "ResultForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tRS As New ADODB.Recordset
Dim I, J, K As Integer
Private Sub Button_Click(Index As Integer)
Select Case Index
Case 5
    Unload Me
End Select
End Sub


Private Sub Command1_Click()
Conn.Execute "delete from printtab"
For I = 1 To 15
Conn.Execute "insert into printtab values(" & I & ",'" & MSF.TextMatrix(I, 0) & "','" & MSF.TextMatrix(I, 1) & "','" & MSF.TextMatrix(I, 3) & "','" & MSF.TextMatrix(I, 4) & "')"
Next

If tRS.State = 1 Then tRS.Close
tRS.Open "select * from printtab", Conn
Set ExamReport1.DataSource = tRS
ExamReport1.Show
End Sub

Private Sub Form_Load()
Dim t1, t2, t3 As Integer
MSF.Cols = 10
MSF.Rows = 40
I = 1
'RegNoVar = 1001
MSF.ColWidth(0) = 2000
MSF.ColWidth(1) = 2000
MSF.ColWidth(2) = 100
MSF.ColWidth(3) = 2000
MSF.ColWidth(4) = 2000
If tRS.State = 1 Then tRS.Close
tRS.Open "select * from StuMas where regno='" & RegNoVar & "'", Conn
If tRS.EOF = False Then

MSF.TextMatrix(I, 0) = "RegNo"
MSF.TextMatrix(I, 1) = RegNoVar
MSF.TextMatrix(I, 3) = "Department"
MSF.TextMatrix(I, 4) = tRS.Fields("dName")

I = I + 1
MSF.TextMatrix(I, 0) = "Name"
MSF.TextMatrix(I, 1) = tRS.Fields("sName")
MSF.TextMatrix(I, 3) = "Year"
MSF.TextMatrix(I, 4) = tRS.Fields("ayear")

I = I + 1
MSF.TextMatrix(I, 0) = "Parent Name"
MSF.TextMatrix(I, 1) = tRS.Fields("pName")
MSF.TextMatrix(I, 3) = "Semister"
MSF.TextMatrix(I, 4) = tRS.Fields("sem")


End If

I = I + 2

MSF.TextMatrix(I, 0) = "Subject Name"

MSF.TextMatrix(I, 3) = "Main Marks"
MSF.TextMatrix(I, 4) = "Total"

MSF.TextMatrix(I, 1) = "Internal Marks"

I = I + 1
If tRS.State = 1 Then tRS.Close
tRS.Open "select subname,eMarks from examMarks where regno='" & RegNoVar & "'", Conn
Do While tRS.EOF = False
MSF.TextMatrix(I, 0) = tRS.Fields("subName")

    If RS.State = 1 Then RS.Close
    RS.Open "select avg(tMarks) from testmarks where regno='" & RegNoVar & "' and subname='" & tRS.Fields("subName") & "'", Conn
    If RS.EOF = False Then
    MSF.TextMatrix(I, 1) = RS(0) & ""
    t1 = t1 + RS(0)
    End If
    
  
    
MSF.TextMatrix(I, 3) = tRS.Fields("eMarks")
    t2 = t2 + tRS.Fields("eMarks")
MSF.TextMatrix(I, 4) = Val(MSF.TextMatrix(I, 1)) + Val(MSF.TextMatrix(I, 3))
    t3 = t3 + Val(MSF.TextMatrix(I, 4))
I = I + 1
tRS.MoveNext
Loop
I = I + 1
MSF.TextMatrix(I, 0) = "Total Marks"
MSF.TextMatrix(I, 1) = t1 & ""
MSF.TextMatrix(I, 3) = t2 & ""
MSF.TextMatrix(I, 4) = t3 & ""


End Sub
