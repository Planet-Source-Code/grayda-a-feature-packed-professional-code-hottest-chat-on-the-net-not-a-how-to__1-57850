VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NChat Options"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6060
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel!!"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK!!"
      Default         =   -1  'True
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   4440
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4215
      Left            =   83
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7435
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "User"
      TabPicture(0)   =   "frmOptions.frx":2CFA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label10"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label11"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "ImageCombo1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Combo1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Combo2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Command6"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "NChat_Button1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "General"
      TabPicture(1)   =   "frmOptions.frx":2D16
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Check1"
      Tab(1).Control(1)=   "Command3"
      Tab(1).Control(2)=   "Check3"
      Tab(1).Control(3)=   "Label2"
      Tab(1).Control(4)=   "Label1"
      Tab(1).Control(5)=   "Label3"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Admin"
      TabPicture(2)   =   "frmOptions.frx":2D32
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label8"
      Tab(2).Control(1)=   "Label7"
      Tab(2).Control(2)=   "Command5"
      Tab(2).Control(3)=   "Command4"
      Tab(2).Control(4)=   "Text3"
      Tab(2).Control(5)=   "Text2"
      Tab(2).Control(6)=   "Text4"
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "Sharing"
      TabPicture(3)   =   "frmOptions.frx":2D4E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label19"
      Tab(3).Control(1)=   "Label18"
      Tab(3).Control(2)=   "Command15"
      Tab(3).Control(3)=   "Command14"
      Tab(3).Control(4)=   "Text12"
      Tab(3).Control(5)=   "Check4"
      Tab(3).Control(6)=   "Dir1"
      Tab(3).Control(7)=   "Text11"
      Tab(3).ControlCount=   8
      Begin VB.CommandButton NChat_Button1 
         Caption         =   "Save ALL UserIcons"
         Height          =   375
         Left            =   3480
         TabIndex        =   33
         Top             =   1560
         Width           =   1695
      End
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -74880
         TabIndex        =   30
         Top             =   1500
         Width           =   5655
      End
      Begin VB.DirListBox Dir1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1665
         Left            =   -74880
         TabIndex        =   29
         Top             =   1860
         Width           =   5655
      End
      Begin VB.CheckBox Check4 
         Appearance      =   0  'Flat
         Caption         =   "Do not share any of my files with everyone"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -74880
         TabIndex        =   28
         Top             =   3780
         Width           =   3375
      End
      Begin VB.TextBox Text12 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -73320
         TabIndex        =   27
         Text            =   "*.*"
         Top             =   540
         Width           =   735
      End
      Begin VB.CommandButton Command14 
         Caption         =   "?"
         Height          =   255
         Left            =   -72480
         TabIndex        =   26
         Top             =   540
         Width           =   375
      End
      Begin VB.CommandButton Command15 
         Caption         =   "View my shared files now"
         Height          =   375
         Left            =   -74760
         TabIndex        =   25
         Top             =   900
         Width           =   2655
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   2055
         Left            =   -74640
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   22
         Text            =   "frmOptions.frx":2D6A
         Top             =   540
         Width           =   4935
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   -73680
         TabIndex        =   21
         Top             =   2700
         Width           =   3975
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   -73680
         TabIndex        =   20
         Top             =   3180
         Width           =   3975
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Check Password"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -71160
         TabIndex        =   19
         Top             =   3660
         Width           =   1455
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Admin Logout"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -72720
         TabIndex        =   18
         Top             =   3660
         Width           =   1335
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Preview"
         Height          =   375
         Left            =   3120
         TabIndex        =   7
         Top             =   2640
         Width           =   1215
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "frmOptions.frx":2EE1
         Left            =   120
         List            =   "frmOptions.frx":2F00
         TabIndex        =   6
         Top             =   2880
         Width           =   2895
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmOptions.frx":2F27
         Left            =   120
         List            =   "frmOptions.frx":2F46
         TabIndex        =   5
         Top             =   2520
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   3255
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "Allow swearing?"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -74880
         TabIndex        =   8
         Top             =   960
         Width           =   1935
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Reset NChat"
         Height          =   375
         Left            =   -74880
         TabIndex        =   10
         Top             =   3600
         Width           =   1455
      End
      Begin VB.CheckBox Check3 
         Appearance      =   0  'Flat
         Caption         =   "Allow Smileys?"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -74880
         TabIndex        =   9
         Top             =   1800
         Width           =   1455
      End
      Begin MSComctlLib.ImageCombo ImageCombo1 
         Height          =   330
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Locked          =   -1  'True
         Text            =   "Please Select an Icon"
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Caption         =   "You have chosen not to share any files. To pick a folder to share, uncheck the box below and select a folder"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   -74760
         TabIndex        =   32
         Top             =   1920
         Width           =   5295
      End
      Begin VB.Label Label19 
         Caption         =   "File types to share:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   31
         Top             =   540
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Username:"
         Height          =   255
         Left            =   -74640
         TabIndex        =   24
         Top             =   2700
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Password:"
         Height          =   255
         Left            =   -74640
         TabIndex        =   23
         Top             =   3180
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "This box allows you to change how messages look. Select or type one in both of the lists, and click PREVIEW"
         Height          =   435
         Left            =   120
         TabIndex        =   17
         Top             =   2040
         Width           =   4650
      End
      Begin VB.Label Label10 
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   3240
         Width           =   5055
      End
      Begin VB.Label Label5 
         Caption         =   "This box allows you to change your username for free!!"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   5055
      End
      Begin VB.Label Label6 
         Caption         =   "In this box, you can select your icon. This wil appear next to your username on the list to the right."
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   1080
         Width           =   5055
      End
      Begin VB.Label Label2 
         Caption         =   "If you are easily offended by foul language, NChat can block some language. To do so, click this box below"
         Height          =   495
         Left            =   -74880
         TabIndex        =   13
         Top             =   480
         Width           =   4935
      End
      Begin VB.Label Label1 
         Caption         =   "If you have a slow computer, or would like to disable pictures (Smileys), then uncheck the box below"
         Height          =   375
         Left            =   -74880
         TabIndex        =   12
         Top             =   1320
         Width           =   3735
      End
      Begin VB.Label Label3 
         Caption         =   "If you want to start again, with NO settings file, NCredits etc, then click this button"
         Height          =   375
         Left            =   -74880
         TabIndex        =   11
         Top             =   3120
         Width           =   3495
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Your old username, just for broadcasting
' so it will say oldusername is now known as username
Dim OldUN As String

Private Sub Check4_Click()
' Uncheck the box, files shared,
' file boxes unhidden

' Check the box, files un shared,
' file boxes hidden
If Check4.Value = 1 Then
Dir1.Visible = False
Text11.Visible = False
Else
Dir1.Visible = True
Text11.Visible = True
End If

End Sub

Private Sub Command14_Click()
' Help
MsgBox "This box allows you to set what kind of files to share. To use this feature, type: '*.' and the extension you want to share (for example: *.mp3 will share all your MP3s, while *.jpg will share all your pictures. You may specify more than one type by seperating them with a semicolon ';'", vbQuestion, "Help with file types"
End Sub

Private Sub Command15_Click()
' Show what files you have shared
frmFileList.Show
End Sub

Private Sub Form_Load()
' Selects the first tab, instead of the IDE
' defined tab
SSTab1.Tab = 0

' Loads your options such
' as swearing and your icon
On Error Resume Next
ImageCombo1.ImageList = frmMain.ImageList1

' Sharing your files?
If ShareMyFiles = False Then
Check4.Value = 1
Dir1.Visible = False
Text11.Visible = False
Else
Check4.Value = 0
Dir1.Visible = True
Text11.Visible = True
End If

' If you are an admin, then you can select all the
' icons, if not, then you need to buy them :)
If frmMain.mnuDev.Visible = False Then
For i = 1 To TotalIcons
ImageCombo1.ComboItems.Add , "K" & i, frmMain.ImageList1.ListImages.Item(i).Key, frmMain.ImageList1.ListImages.Item(i).Key
Next i
Else
For i = 1 To frmMain.ImageList1.ListImages.Count - 2
' Prefix the key with a "K" because keys can't be numbers only
ImageCombo1.ComboItems.Add , "K" & i, frmMain.ImageList1.ListImages.Item(i).Key, frmMain.ImageList1.ListImages.Item(i).Key
Next i
End If
' Selects your current icon
ImageCombo1.ComboItems.Item(MyIcon).Selected = True

' If you are an admin, then show the logout button
If frmMain.mnuDev.Visible = True Then Command5.Enabled = True

Text1.Text = UserName
' Old UN lets NChat send messages like: OldUsername
' is now known as NewUsername and stuff
OldUN = UserName
ImageCombo1.ImageList = frmMain.ImageList1
Combo1.Text = StartMSG
Combo2.Text = EndMSG

If Swearing = True Then Check1.Value = Checked
If Smiley = True Then Check3.Value = 1
If Popup = True Then Check2.Value = 1
Text11.Text = frmFileList.File1.Path

Text12.Text = frmFileList.File1.Pattern

End Sub

Private Sub Command1_Click()
' If the checks are enabled, then
' turn the booleans on
On Error GoTo Errors
' CHecks your name for invalid stuff
If Check4.Value = 1 Then
ShareMyFiles = False
Else
ShareMyFiles = True
End If

If Right(Trim(Text1.Text), 3) = "[A]" And frmMain.mnuDev.Visible = False Or Right(Trim(Text1.Text), 4) = "[RA]" And frmMain.mnuDev.Visible = False Then
MsgBox "Sorry, but only administrators can have [A] or [RA] on the end of their name...", vbExclamation, "Bad Username"
SSTab1.Tab = 0
Exit Sub
End If

For i = 1 To frmMain.List1.ListItems.Count
If Trim(frmMain.List1.ListItems.Item(i).Text) = Trim(Text1.Text) And Trim(frmMain.List1.ListItems.Item(i).Text) <> UserName Then
MsgBox "Sorry, that username has been taken. Please select another one...", vbExclamation, "Username in use"
SSTab1.Tab = 0
Exit Sub
End If
Next i


If Len(Text1.Text) > 25 Then
MsgBox "Your username is too long! Please shorten it (25 letters max)", vbCritical, "Username too long!"
SSTab1.Tab = 0
Exit Sub
End If
OldUsername = UserName
UserName = Trim(Text1.Text)

If UserName = OldUsername Then
UserName = OldUsername
Else


Broadcast "chuø" & OldUN & "ø" & UserName & "ø" & MyIcon

DoEvents
' Server Text syntax: svr <delimiter> Text to send
Broadcast "svrø" & OldUsername & " is now known as " & UserName
Text "Changed your Username from " & OldUN & " to " & UserName & "!" & vbCrLf, Heading, True, , , , vbCenter
End If

frmFileList.File1.Path = Text11.Text


If Check1.Value = 1 Then
Swearing = True
Else
Swearing = False
End If


StartMSG = Combo1.Text
EndMSG = Combo2.Text

' Sets MyIcon to imagecombo's icon
If ImageCombo1.SelectedItem.Index <> MyIcon Then
Broadcast "chiø" & UserName & "ø" & ImageCombo1.SelectedItem.Index
MyIcon = ImageCombo1.SelectedItem.Index
frmMain.List1.ListItems(1).SmallIcon = frmMain.ImageList1.ListImages(MyIcon).Key
End If

If Check3.Value = Checked Then
Smiley = True
Else
Smiley = False
End If



frmFileList.File1.Pattern = Text12.Text
FilePattern = Text12.Text
Unload Me
Exit Sub
Errors:
MyIcon = MyIcon
Unload Me
End Sub


Private Sub Command2_Click()
' Cancel Button
Unload Me
End Sub

Private Sub Command3_Click()
' Resets NChat so you can start again
If MsgBox("YOU ARE ABOUT TO DELETE ALL YOUR NCHAT SETTINGS. ARE YOU SURE YOU WANT TO DO THIS?", vbCritical + vbYesNo, "WARNING:") = vbYes Then
Kill AppPath & GetUserName & ".ncg"
MsgBox "NChat settings files deleted. NChat will now reset so you can start again", vbInformation, "Delete successful"
End
Else
MsgBox "Delete aborted"
End If
End Sub

Private Sub Command4_Click()

If Text2.Text = UserName & "101" Then

Enc = ""

For i = 1 To Len(Day(Date) & UserName & Hour(Time) & Minute(Time) & GetUserName)
Enc = Enc & Hex(Asc(Mid(Day(Date) & UserName & Hour(Time) & Minute(Time) & GetUserName, i, 1)) / 2)
Next i

If Text3.Text = Enc Then
frmMain.mnuDev.Visible = True
TrueAdmin = True
Text2.Text = ""
Text3.Text = ""
MsgBox "Password Correct!!", vbInformation, "Administrator Login"
Broadcast "svrø" & UserName & " is now an NChat Admin!"
OLD = UserName
UserName = Replace(UserName, " [A]", "")
UserName = UserName & " [A]"
Broadcast "chuø" & OLD & "ø" & UserName & "ø" & MyIcon
Unload Me
Else
MsgBox "Incorrect Password!!", vbCritical, "Wrong"
Text2.Text = ""
Text3.Text = ""
Exit Sub
End If
Else
MsgBox "Incorrect Username!!", vbCritical, "Wrong"
Text2.Text = ""
Text3.Text = ""
Exit Sub
End If
End Sub

Private Sub Command5_Click()
frmMain.mnuDev.Visible = False
TrueAdmin = False
OLD = UserName
UserName = Replace(UserName, " [A]", "")
Broadcast "chuø" & OLD & "ø" & UserName & "ø" & MyIcon
Command5.Enabled = False
Broadcast "svrø+username+ is no longer an admin!"
End Sub

Private Sub Command6_Click()
Label10.Caption = Combo1.Text & " " & UserName & " " & Combo2.Text & " Hello!"
End Sub


Private Sub Dir1_Click()
Text11.Text = Dir1.List(Dir1.ListIndex)

End Sub




Private Sub NChat_Button1_Click()
For i = 1 To frmMain.ImageList1.ListImages.Count
SavePicture frmMain.ImageList1.ListImages(i).Picture, App.Path & "\UserImages\UserImage" & i & ".ico"
Next i
End Sub

Private Sub Text11_Change()
On Error Resume Next
Dir1.Path = Text11.Text

End Sub

Private Sub Text2_Change()
If Text2.Text > "" And Text3.Text > "" Then
Command4.Enabled = True
Else
Command4.Enabled = False
End If

End Sub

Private Sub Text3_Change()
If Text3.Text > "" And Text2.Text > "" Then
Command4.Enabled = True
Else
Command4.Enabled = False
End If

End Sub

Private Sub UserControl11_Click()

End Sub
