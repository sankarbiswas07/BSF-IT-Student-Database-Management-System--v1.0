VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H80000010&
   Caption         =   "Student Database Administration"
   ClientHeight    =   3690
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   12390
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu MastentMenu 
      Caption         =   "Master Entries"
      Begin VB.Menu DeptMenu 
         Caption         =   "Department Details"
         Shortcut        =   {F1}
      End
      Begin VB.Menu SubDetMenu 
         Caption         =   "Subject Details"
         Shortcut        =   {F2}
      End
   End
   Begin VB.Menu StudentRegMenu 
      Caption         =   "Student Registration"
   End
   Begin VB.Menu AttEntryMenu 
      Caption         =   "Attendance Entry"
   End
   Begin VB.Menu IntMarksEntryMenu 
      Caption         =   "Int Marks Entry"
   End
   Begin VB.Menu MainExamMEnu 
      Caption         =   "Main Exam Marks Entry"
   End
   Begin VB.Menu ReportMenu 
      Caption         =   "Report"
   End
   Begin VB.Menu apps 
      Caption         =   "Application"
      Begin VB.Menu cald 
         Caption         =   "Calender"
         Shortcut        =   {F3}
      End
      Begin VB.Menu calc 
         Caption         =   "Calculator"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu abt 
      Caption         =   "About"
      Begin VB.Menu crdt 
         Caption         =   "Credit"
         Shortcut        =   {F5}
      End
      Begin VB.Menu hlp 
         Caption         =   "Help"
         Shortcut        =   {F6}
      End
   End
   Begin VB.Menu Exitmenu 
      Caption         =   "Exit"
      Begin VB.Menu lgot 
         Caption         =   "Logout"
         Shortcut        =   {F7}
      End
      Begin VB.Menu cls 
         Caption         =   "Close"
         Shortcut        =   {F8}
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AttEntryMenu_Click()
AttEntryForm.Show
End Sub

Private Sub calc_Click()
Call Shell(App.Path & "\calc.exe", vbNormalFocus)
End Sub

Private Sub cald_Click()
calenderfrm.Show
End Sub

Private Sub cls_Click()
End
End Sub

Private Sub crdt_Click()
poster.Show
End Sub

Private Sub DeptMenu_Click()
DepForm.Show
End Sub


Private Sub hlp_Click()
help.Show
End Sub

Private Sub IntMarksEntryMenu_Click()
TestEForm.Show
End Sub

Private Sub lgot_Click()
Unload Me
LoginForm.Show
End Sub

Private Sub MainExamMEnu_Click()
ExamEForm.Show
End Sub

Private Sub Label8_Click()
ReportForm.Show
End Sub

Private Sub ReportMenu_Click()
ReportForm.Show
End Sub

Private Sub StudentRegMenu_Click()
StuEntryForm.Show
End Sub

Private Sub SubDetMenu_Click()
SubForm.Show
End Sub
