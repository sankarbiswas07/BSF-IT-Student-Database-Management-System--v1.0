VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmopen 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7725
   ClientLeft      =   5760
   ClientTop       =   3390
   ClientWidth     =   9495
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   9495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MCI.MMControl MMControl2 
      Height          =   330
      Left            =   9960
      TabIndex        =   9
      Top             =   10440
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   582
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   "C:\WINDOWS\Desktop\exam\bit.avi"
   End
   Begin VB.Timer Timer3 
      Left            =   360
      Top             =   4440
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   420
      Left            =   5760
      TabIndex        =   2
      Top             =   3120
      Visible         =   0   'False
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   741
      _Version        =   327682
      Appearance      =   1
   End
   Begin MCI.MMControl MMControl1 
      Height          =   330
      Left            =   5880
      TabIndex        =   6
      Top             =   10440
      Visible         =   0   'False
      Width           =   3150
      _ExtentX        =   5556
      _ExtentY        =   582
      _Version        =   393216
      PlayVisible     =   0   'False
      DeviceType      =   ""
      FileName        =   "C:\WINDOWS\Desktop\exam\music.mid"
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   840
      Picture         =   "frmopen.frx":0000
      ScaleHeight     =   2955
      ScaleWidth      =   4155
      TabIndex        =   5
      Top             =   7560
      Width           =   4215
   End
   Begin VB.Timer Timer2 
      Left            =   360
      Top             =   3480
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit !"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   3
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Click Here to Enter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   1
      Top             =   1560
      Width           =   3495
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   360
      Top             =   2400
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Height          =   405
      Left            =   2880
      TabIndex        =   8
      Top             =   2640
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   714
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Text            =   "Panel"
            TextSave        =   "Panel"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   9022
            Text            =   "Progress Bar"
            TextSave        =   "Progress Bar"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            TextSave        =   "10:25 PM"
            Key             =   "ProgBar"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      Height          =   2535
      Left            =   5160
      Shape           =   2  'Oval
      Top             =   5040
      Width           =   4215
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000009&
      Caption         =   "     Developed                  By                  Gautam P                 and                  Arun S N"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1815
      Left            =   6000
      TabIndex        =   4
      Top             =   5400
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Welcome to Attendance and Examination Database"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   495
      Left            =   5280
      TabIndex        =   0
      Top             =   720
      Width           =   6915
   End
End
Attribute VB_Name = "frmopen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' This form demonstrates how to place a Progress-Bar into a
' panel of a status bar.
'

Private Sub Command1_Click()
Unload Me
frmlogin.Show
End Sub

Private Sub ShowProgressInStatusBar(ByVal bShowProgressBar As Boolean)

    Dim tRC As RECT
    
    If bShowProgressBar Then

' Get  size of the Panel
' (2) Rectangle from the status bar
' remember that Indexes in the API are always 0 based (well,
' nearly always) - therefore Panel(2) = Panel(1) to the api
'
'
        SendMessageAny StatusBar1.hwnd, SB_GETRECT, 1, tRC
'
' and convert it to twips....
'
        With tRC
            .Top = (.Top * Screen.TwipsPerPixelY)
            .Left = (.Left * Screen.TwipsPerPixelX)
            .Bottom = (.Bottom * Screen.TwipsPerPixelY) - .Top
            .Right = (.Right * Screen.TwipsPerPixelX) - .Left
        End With
'
' Now Reparent the ProgressBar to the statusbar
'
        With ProgressBar1
            SetParent .hwnd, StatusBar1.hwnd
            .Move tRC.Left, tRC.Top, tRC.Right, tRC.Bottom
            .Visible = True
            .Value = 0
        End With
        
    Else
'
' Reparent the progress bar back to the form and hide it
'
        SetParent ProgressBar1.hwnd, Me.hwnd
        ProgressBar1.Visible = False
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
'
' Should really re-parent the progress bar here,
' just in case anything wrong happened
'
    ShowProgressInStatusBar False
    
End Sub

Private Sub Timer1_Timer()
'
' This timer routine simply updates the progress bar to make it

    Static lCount As Long
    
    lCount = lCount + 5
    
    If lCount > 100 Then
        Timer1.Enabled = False
        ShowProgressInStatusBar False
        Command1.Enabled = True
        ProgressBar1.Visible = False
        StatusBar1.Visible = False
        lCount = 0
    End If
    
    ProgressBar1.Value = lCount
    
End Sub


Private Sub cmdexit_Click()
RestoreLastCursor
End
End Sub

Private Sub Form_Load()
MMControl2.Notify = False
MMControl2.Wait = True
MMControl2.Shareable = False
MMControl2.DeviceType = "AVIVideo"
MMControl2.Command = "Open"
MMControl2.Command = "Play"
MMControl2.Wait = True
MMControl2.Command = "Play"
MMControl2.Command = "Play"
MMControl2.Command = "Stop"
MMControl2.Command = "Close"
'
' Disable this button for now
'
    Command1.Enabled = False
'
' Setup the progress bar with some values
'
    ProgressBar1.Min = 0
    ProgressBar1.Max = 100
'
' Show ProgressBar in Status Bar
'
    ShowProgressInStatusBar True
'
' Enable the timer so it looks like we're doing something
'
Timer1.Enabled = True
StartAnimatedCursor (App.Path & "\MYSTER~1.ani")
MMControl1.Notify = False
MMControl1.Wait = True
MMControl1.Shareable = False
MMControl1.DeviceType = "Sequencer"
MMControl1.Command = "Open"
MMControl1.Command = "Play"
Z = Label1.Left
Timer3.Interval = 100
Timer2.Interval = 500
Call opencon
End Sub

Private Sub Timer3_Timer()
If Not Z < -6720 Then
  Z = Z - 200
  Label1.Left = Z
 If Z < -1600 Then
     Z = 15640
     Label1.Left = Z
 End If
End If
End Sub


Private Sub Timer2_Timer()
If Label3.Visible = True Then
Label3.Visible = False
Else
Label3.Visible = True
End If
End Sub
