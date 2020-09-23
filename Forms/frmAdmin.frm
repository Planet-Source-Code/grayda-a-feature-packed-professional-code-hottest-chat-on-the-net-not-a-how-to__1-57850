VERSION 5.00
Begin VB.Form frmAdmin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Administrators Toolbox"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3690
   ControlBox      =   0   'False
   Icon            =   "frmAdmin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   3690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Close this window"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   3495
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Send Server (Purple) Messages"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Welcome Message"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      ToolTipText     =   "Set the message that everyone will see on entry"
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Kick User"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Kick someone from the NChat room"
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmAdmin is a small toolbox that allows people who
' have created rooms (Room Admins), to administer their
' room. They have limited power, but is enough


Private Sub Command1_Click()
' Showbox brings up a box with all users in the room in it
' So we can pick a name, and click it, without having to
' copy a username, paste it and click OK
ShowBox "Kick User", "Kick a user from NChat"
Broadcast "ksvø" & SelUser & "ø+username+"
End Sub

Private Sub Command2_Click()
' Your welcome message

' If no message is set, then say: message for this room SET
' but if the message is being updated, then say so
If WelcomeMsg = "" Then
WelcomeMsg = "svr+d+" & InputBox("Enter a welcome message that everyone will see", "Welcome Message", WelcomeMsg)
Broadcast "svrøWelcome message for this room set!"
End If
End Sub

Private Sub Command5_Click()
' Sends purple (Server) messages
Broadcast "svrø" & InputBox("Enter text to send as purple", "Send Server Text")
End Sub

Private Sub Command6_Click()
' Give some warning about closing the box
If MsgBox("Warning. Once you close this box, then you cannot re-open it unless you create a new room. Are you sure you want to close this window?", vbQuestion + vbYesNo, "Close this window?") = vbYes Then Unload Me
End Sub
