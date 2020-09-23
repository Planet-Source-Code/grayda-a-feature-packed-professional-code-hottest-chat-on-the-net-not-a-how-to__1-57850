VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAutoBotOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "The Notch Control Panel"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   7050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   2775
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   4895
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Idle Chat Options"
      TabPicture(0)   =   "frmAutoBotOptions.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Timer1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Check1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Phrase Learning"
      TabPicture(1)   =   "frmAutoBotOptions.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label5"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Timer2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "NChat_Button1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "NChat_Button2"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      Begin VB.CommandButton NChat_Button2 
         Caption         =   "Start Learning"
         Height          =   495
         Left            =   -74880
         TabIndex        =   7
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton NChat_Button1 
         Caption         =   "Stop Learning"
         Height          =   495
         Left            =   -73560
         TabIndex        =   6
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Timer Timer2 
         Interval        =   1000
         Left            =   -68880
         Top             =   2160
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Text            =   "10"
         Top             =   1800
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Allow Idle Chatter?"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   30000
         Left            =   5880
         Top             =   1800
      End
      Begin VB.Label Label5 
         Caption         =   $"frmAutoBotOptions.frx":0038
         Height          =   855
         Left            =   -74880
         TabIndex        =   10
         Top             =   480
         Width           =   6255
      End
      Begin VB.Label Label4 
         Caption         =   $"frmAutoBotOptions.frx":0115
         Height          =   855
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   6255
      End
      Begin VB.Label Label1 
         Caption         =   "New Phrases added to DB: 0"
         Height          =   255
         Left            =   -74880
         TabIndex        =   8
         Top             =   2160
         Width           =   3255
      End
      Begin VB.Label Label3 
         Caption         =   "Seconds"
         Height          =   255
         Left            =   1920
         TabIndex        =   5
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Interval:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1800
         Width           =   615
      End
   End
   Begin VB.CommandButton NChat_Button4 
      Caption         =   "Hide Window"
      Height          =   375
      Left            =   5040
      TabIndex        =   0
      Top             =   3000
      Width           =   1815
   End
End
Attribute VB_Name = "frmAutoBotOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
' Timer1 handles our idle chatter,
' When Timer1_Timer is triggered,
' then Notch will speak a random phrase
If Check1.Value = 1 Then
Timer1.Enabled = True
frmAutobot.List1.AddItem "(" & Time & ") Notch will now start speaking random phrases every " & Timer1.Interval / 10000 & " seconds."
Else
Timer1.Enabled = False
frmAutobot.List1.AddItem "(" & Time & ") Notch has stopped speaking random phrases"
End If


End Sub



Private Sub Command1_Click()
' Opens the dialog box to let us select a script
D1.Filter = "Notch INI File (*.ini)|*.ini|All Files (*.*)|*.*"
D1.ShowOpen

If D1.filename > "" Then
List1.AddItem "(" & Time & ") Notch Exclusion File Loaded!"
Text1.Text = D1.filename

End Sub

Private Sub Command2_Click()
If Trim(Text1.Text) > "" Then
IniFile(3) = Text1.Text
Else
MsgBox "No exlude file loaded! Please load one now!", vbCritical, "No file loaded!"
Exit Sub
End If

ContextualSentences = True

frmAutobot.List1.AddItem "(" & Time & ") Notch has contextual phrasing on, with a randomness of " & Text3.Text & "%"

End Sub

Private Sub Command3_Click()
frmAutobot.List1.AddItem "(" & Time & ") Notch has contextual phrasing turned OFF"

ContextualSentences = False
End Sub

Private Sub Form_Load()
SSTab1.Tab = 0
End Sub

Private Sub NChat_Button1_Click()
NotchLearning = False
frmAutobot.List1.AddItem "(" & Time & ") Notch has stopped learning new phrases. He learnt " & NewWords & " new phrases!"

End Sub

Private Sub NChat_Button2_Click()
NotchLearning = True
frmAutobot.List1.AddItem "(" & Time & ") Notch is currently learning new phrases"


End Sub

Private Sub NChat_Button4_Click()
Me.Visible = False

End Sub

Private Sub Timer1_Timer()
' This is our Idle Phrase broadcaster
Dim Sects As Integer
If IniFile(2) = "" Then Exit Sub
Sects = 1
Do Until ReadText("IdlePhrase", "Phrase" & Sects, 2) = ""
Sects = Sects + 1
Loop
' Ensures 0 doesn't come up
R = Int(Rnd * Sects) + 1

Broadcast ReadText("IdlePhrase", "Phrase" & R, 2)

End Sub

Private Sub Timer2_Timer()
Label1.Caption = "New Phrases added to DB: " & NewWords
End Sub
