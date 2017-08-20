VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FlashScreen 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   5550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7470
   LinkTopic       =   "Form2"
   ScaleHeight     =   5550
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3960
      Top             =   3960
   End
   Begin VB.PictureBox Picture1 
      Height          =   5535
      Left            =   0
      Picture         =   "FlashScreen.frx":0000
      ScaleHeight     =   5475
      ScaleWidth      =   7395
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   135
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   495
      End
   End
End
Attribute VB_Name = "FlashScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
I = 1
'Conn.Open "Provider=MSDAORA.1;User ID=scott;password=tiger;Persist Security Info=true"   'oracle
If Conn.State = 1 Then Conn.Close
Conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\sData.mdb;Persist Security Info=False" 'msaccess
'Conn.Open "Provider=SQLOLEDB.1;Persist Security Info=False;Initial Catalog=maindata;Data Source=SQL1"
ProgressBar1.Value = ProgressBar1.Min
End Sub


Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 1
If ProgressBar1.Value = 50 Then
ProgressBar1.Value = ProgressBar1 + 50
LoginForm.Show
If ProgressBar1.Value >= ProgressBar1.Max Then
Unload Me
Timer1.Enabled = False
End If
End If
End Sub



