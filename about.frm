VERSION 5.00
Begin VB.Form poster 
   Caption         =   "POSTER"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   ScaleHeight     =   6810
   ScaleWidth      =   9465
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   6975
      Left            =   -120
      Picture         =   "about.frx":0000
      ScaleHeight     =   6915
      ScaleWidth      =   9555
      TabIndex        =   0
      Top             =   -120
      Width           =   9615
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   6360
         TabIndex        =   1
         Top             =   6480
         Width           =   3015
      End
   End
End
Attribute VB_Name = "poster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()
Unload Me
End Sub
