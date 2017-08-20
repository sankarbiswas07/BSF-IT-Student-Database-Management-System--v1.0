VERSION 5.00
Begin VB.Form help 
   Caption         =   "Help...!!!"
   ClientHeight    =   1875
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   ScaleHeight     =   1875
   ScaleWidth      =   5385
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   1935
      Left            =   0
      Picture         =   "help.frx":0000
      ScaleHeight     =   1875
      ScaleWidth      =   5355
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   1095
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   1095
      End
   End
End
Attribute VB_Name = "help"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()
Unload Me
End Sub
