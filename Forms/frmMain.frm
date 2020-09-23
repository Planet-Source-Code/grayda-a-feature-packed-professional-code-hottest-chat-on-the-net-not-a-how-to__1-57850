VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Welcome to NChat Alpha - Waiting for incoming connections..."
   ClientHeight    =   6330
   ClientLeft      =   165
   ClientTop       =   495
   ClientWidth     =   8340
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   8340
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView ListView1 
      Height          =   1005
      Left            =   6360
      TabIndex        =   6
      Top             =   4560
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1773
      View            =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
      Picture         =   "frmMain.frx":2CFA
   End
   Begin VB.ComboBox txtSend 
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   5640
      Width           =   6135
   End
   Begin MSComctlLib.ListView lstIgnore 
      Height          =   1005
      Left            =   6360
      TabIndex        =   4
      Top             =   4560
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1773
      View            =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
      Picture         =   "frmMain.frx":3301
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "/action"
      Height          =   255
      Left            =   7440
      TabIndex        =   3
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send!!"
      Height          =   255
      Left            =   6360
      TabIndex        =   2
      Top             =   5640
      Width           =   975
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2520
      Tag             =   " "
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   39
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3863
            Key             =   "Default"
            Object.Tag             =   "People"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3BFD
            Key             =   "Devil"
            Object.Tag             =   "Evil"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3F97
            Key             =   "Agent Smith"
            Object.Tag             =   "Evil"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4331
            Key             =   "Radioactive"
            Object.Tag             =   "Evil"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":46CB
            Key             =   "Smiley"
            Object.Tag             =   "People"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4A65
            Key             =   "NChat"
            Object.Tag             =   "Misc"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":776F
            Key             =   "Star"
            Object.Tag             =   "Misc"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7B09
            Key             =   "Lightning"
            Object.Tag             =   "Misc"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7EA3
            Key             =   "Half-Life"
            Object.Tag             =   "Games"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":823D
            Key             =   "One Fingered Salute"
            Object.Tag             =   "Misc"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":85D7
            Key             =   "GTAIII"
            Object.Tag             =   "Games"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8971
            Key             =   "Boy n Girl"
            Object.Tag             =   "People"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8D24
            Key             =   "Chinese"
            Object.Tag             =   "People"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":92F2
            Key             =   "The Finger"
            Object.Tag             =   "Misc"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":97F4
            Key             =   "Idiot"
            Object.Tag             =   "People"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9BD9
            Key             =   "Weed"
            Object.Tag             =   "Misc"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A075
            Key             =   "Power"
            Object.Tag             =   "Misc"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A40F
            Key             =   "Play"
            Object.Tag             =   "Misc"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A7A9
            Key             =   "MOHA"
            Object.Tag             =   "Games"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AB43
            Key             =   "Girl"
            Object.Tag             =   "People"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AEF5
            Key             =   "Boy"
            Object.Tag             =   "People"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B2A7
            Key             =   "Rammstein 1"
            Object.Tag             =   "Music"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B841
            Key             =   "Rammstein 2"
            Object.Tag             =   "Music"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BBDB
            Key             =   "Green Eye"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C175
            Key             =   "Evanescence"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C70F
            Key             =   "Evil 1"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":CFE9
            Key             =   "Evil 2"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D583
            Key             =   "Nemo"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":DB1D
            Key             =   "Intruder"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E3F7
            Key             =   "Gold Star"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":ECD1
            Key             =   "Delta Force"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F5AB
            Key             =   "Outkast"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FE85
            Key             =   "Ozzy Osbourne"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1075F
            Key             =   "Gun"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11015
            Key             =   "Halo"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11CEF
            Key             =   "Alert!!"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":129C9
            Key             =   "Mail"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13A1B
            Key             =   "Unavailable"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13DB5
            Key             =   "AFK"
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock sckRooms 
      Left            =   2040
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemoteHost      =   "255.255.255.255"
      RemotePort      =   2222
      LocalPort       =   2222
   End
   Begin MSComctlLib.StatusBar SB1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   1
      Top             =   6060
      Width           =   8340
      _ExtentX        =   14711
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3016
            MinWidth        =   882
            Text            =   "NChat - Disconnected!"
            TextSave        =   "NChat - Disconnected!"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10530
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   582
            MinWidth        =   423
            Picture         =   "frmMain.frx":1414F
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   2040
      Top             =   720
   End
   Begin MSComDlg.CommonDialog dlgSave 
      Left            =   1440
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Save NChat Log"
      Filter          =   "Text Files (*.TXT)|*.txt|RTF File (*.rtf)|*.RTF|All Files (*.*)|*.*"
   End
   Begin MSWinsockLib.Winsock sckUDP 
      Left            =   1440
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemoteHost      =   "255.255.255.255"
      RemotePort      =   1113
      LocalPort       =   1113
   End
   Begin MSComctlLib.ListView List1 
      Height          =   4335
      Left            =   6360
      TabIndex        =   0
      Top             =   120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   7646
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FlatScrollBar   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Username"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Files"
         Object.Width           =   917
      EndProperty
      Picture         =   "frmMain.frx":144A1
   End
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   5415
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   9551
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmMain.frx":14964
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image picGreen 
      Height          =   240
      Left            =   0
      Picture         =   "frmMain.frx":149E2
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Menu mnuFile 
      Caption         =   "Main Menu"
      Begin VB.Menu mnuConnection 
         Caption         =   "Connection"
         Begin VB.Menu mnuConnect 
            Caption         =   "Connect to NChat Server"
         End
         Begin VB.Menu hrCon 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDisconnect 
            Caption         =   "Disconnect from Server"
         End
      End
      Begin VB.Menu MMHR 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNChatProf 
         Caption         =   "NChat Profiles"
         Begin VB.Menu mnuNewProfile 
            Caption         =   "New Profile"
         End
         Begin VB.Menu mnuLoadProfile 
            Caption         =   "Load Profile"
            Shortcut        =   ^L
         End
      End
      Begin VB.Menu hr1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSend 
         Caption         =   "Send a file"
         Visible         =   0   'False
      End
      Begin VB.Menu hr2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSaveChat 
         Caption         =   "Save Chat Log"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "Clear Chat Text"
      End
      Begin VB.Menu hr3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "NChat Options"
         Shortcut        =   ^O
      End
      Begin VB.Menu hr4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "Quit NChat"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuNCredits 
      Caption         =   "NCredits"
      Begin VB.Menu mnuBalance 
         Caption         =   "NCredits Balance"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuStore 
         Caption         =   "NChat Store"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      Begin VB.Menu mnuPM 
         Caption         =   "Private Messages"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuIgnore2 
         Caption         =   "Ignore List"
      End
   End
   Begin VB.Menu mnuChatRooms 
      Caption         =   "Chat Rooms"
      Begin VB.Menu mnuJoinCustom 
         Caption         =   "List all rooms"
      End
      Begin VB.Menu mnuCreateRoom 
         Caption         =   "Create your own room"
      End
      Begin VB.Menu hr5 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuChangeRoomInfo 
         Caption         =   "Change room info"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuDev 
      Caption         =   "Admin Menu"
      Begin VB.Menu mnuAdminWindow 
         Caption         =   "NChat Server Log"
         Shortcut        =   ^{INSERT}
      End
      Begin VB.Menu mnuBL 
         Caption         =   "Broadcast / Loopback"
      End
      Begin VB.Menu mnuRawData 
         Caption         =   "Raw Data"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuAutoBot 
         Caption         =   "Start / Stop Bot"
      End
      Begin VB.Menu mnuChangeNCredits 
         Caption         =   "Change your NCredits"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuSendPic 
         Caption         =   "Send RTF Picture"
      End
      Begin VB.Menu mnuChatGoodies 
         Caption         =   "Chatroom Goodies"
         Begin VB.Menu mnuFake 
            Caption         =   "Fake Users"
            Begin VB.Menu mnuInsertFake 
               Caption         =   "Insert Fake User"
               Shortcut        =   ^I
            End
            Begin VB.Menu mnuRemFake 
               Caption         =   "Remove Fake User"
               Shortcut        =   ^D
            End
         End
         Begin VB.Menu mnuWmsg 
            Caption         =   "Welcome Message"
            Shortcut        =   ^W
         End
         Begin VB.Menu mnuDoHeading 
            Caption         =   "Create a heading"
            Shortcut        =   ^H
         End
         Begin VB.Menu mnuSendAll 
            Caption         =   "Send a message to all rooms"
         End
         Begin VB.Menu mnuNewRoom 
            Caption         =   "New Room with set ID"
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      WindowList      =   -1  'True
      Begin VB.Menu mnuTextHelp 
         Caption         =   "Display Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuSmileys 
         Caption         =   "Smileys you can use"
      End
      Begin VB.Menu mnuTipofTheDay 
         Caption         =   "Tip of the day"
      End
      Begin VB.Menu hrA 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About NChat"
      End
   End
   Begin VB.Menu mnuUserList 
      Caption         =   "User List"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu ul1 
         Caption         =   "Send >> a Private Message"
      End
      Begin VB.Menu ul2 
         Caption         =   "Send >> some NCredits"
      End
      Begin VB.Menu mnuUEIgnore 
         Caption         =   "Ignore >>"
      End
      Begin VB.Menu ul4 
         Caption         =   "Browse >>'s Files"
      End
      Begin VB.Menu mnuUE 
         Caption         =   "Admin"
         Visible         =   0   'False
         Begin VB.Menu mnuUEGhost 
            Caption         =   "Ghost"
         End
         Begin VB.Menu mnuUEKick 
            Caption         =   "Kick"
         End
         Begin VB.Menu mnuUEAdmin 
            Caption         =   "Make Admin"
         End
         Begin VB.Menu mnuUERemAdmin 
            Caption         =   "Kill Admin"
         End
         Begin VB.Menu mnuUERedirect 
            Caption         =   "Redirect"
         End
         Begin VB.Menu mnuUEPIP 
            Caption         =   "Print Info"
         End
         Begin VB.Menu hr6 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRealUser 
            Caption         =   "Is this user real?"
         End
      End
   End
   Begin VB.Menu mnuInfo 
      Caption         =   "Information Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuAvailable 
         Caption         =   "I am Available for chat"
      End
      Begin VB.Menu mnuAFK 
         Caption         =   "I am Away from keyboard"
      End
      Begin VB.Menu mnuUnAvailable 
         Caption         =   "I am Unavailable for chat"
      End
      Begin VB.Menu hr7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAwayMSG 
         Caption         =   "Set Custom Away Message"
      End
      Begin VB.Menu hr8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIgnore 
         Caption         =   "Ignore List"
      End
      Begin VB.Menu hr9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUECHI 
         Caption         =   "Change my User Icon"
      End
      Begin VB.Menu mnuUECHU 
         Caption         =   "Change my Username"
      End
   End
   Begin VB.Menu mnuTray 
      Caption         =   "NChat Tray Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuTray_Show 
         Caption         =   "Show NChat"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHide 
         Caption         =   "Completely Hide NChat"
      End
      Begin VB.Menu hr10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTray_Available 
         Caption         =   "I am Available for chat"
      End
      Begin VB.Menu mnuTray_AFK 
         Caption         =   "I am Away From Keyboard"
      End
      Begin VB.Menu mnuTray_Unavailable 
         Caption         =   "I am Unavailable for chat"
      End
      Begin VB.Menu hr11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTray_Quit 
         Caption         =   "Quit NChat"
      End
   End
   Begin VB.Menu mnuPopup_Ignore 
      Caption         =   "Ignore Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuDeleteIgnore 
         Caption         =   "Delete Ignore"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' FireStorm Visual Software in conjunction with
' Solid Inc. Represents:

' NChat Alpha Build 10 Series
'   Coded and Researched by Grayda
'       firestorm_visual@hotmail.com

' This is one hell of a gift for the people at
' planetsourcecode.com, since most of this code
' came from here :P. This code has taken me a year
' to write, doing it in my free time (Which wasn't
' very often :) )

' You may use any parts of this code, as long as FVS is
' mentioned in the code / web-site / program etc.
' Check out the NEW License.txt for more info!

' This code is hopefully free of bugs, loopholes
' and assorted coding errors. If you come across
' any more, then contact me at: firestorm_visual@hotmail.com

' Before we begin, I guess some
' clarification is needed:


' In this chat program, you will recieve any and all
' messages that you send. You may be thinking 'why?'.
' It's simple really: Recieving your own messages
' does absolutely no harm, so why write extra
' code just to stop it?

' This code was written from the ground up, about
' 12 months ago. Since then, I have worked hard on
' this. And most of the code was new to me, since
' this is my second attempt at a UDP Chat, so
' please forgive it's poor structure, bad coding
' and control name inconsistency

' This code allows incoming data to be case insensitive,
' meaning that USERNAME and uSeRnAmE are the same
Option Compare Text

' Your Away Message, that is sent to everyone
' who tries to contact you via Private Messages
' when your status is set to anything other than
' available (ie. AFK, Unavailable)
Dim AwayMSG As String

' OldData was the last message to be recieved through
' the winsock. It exists, so people cannot rack
' up NCredits by spamming the room. You can still
' send lots of messages, but you won't recieve
' NCredits for them
Dim OldData As String

' If Ban = True, then when you exit NChat,
' you are banned forever, or until you
' delete your settings file and start again
Dim Ban As Boolean

Private Sub cmdAction_Click()
' An action lets you act out certain things.
' the action is replaced by your username,
' so /action screams becomes <USERNAME> screams
txtSend.Text = "/action " & txtSend.Text
End Sub

Private Sub cmdSend_Click()
' The Send button at the bottom right of frmMain

' This sub has been completely re-written, because
' it had a small security risk in there. People
' could get 1 NCredits each time for flooding the
' room with /action messages

' Stop Empty-Messages from being sent
If Trim(txtSend.Text) = "" Then Exit Sub
' This code ensures that you can't raise errors
' by trying to chat on an 'unconnected' sock
If SB1.Panels(1).Text = "NChat - Disconnected!" Then
    MsgBox "There was an error sending your message. It appears that you are not connected to the NChat server. Check the little light in the corner. If it is red, then please wait 30 seconds, or restart NChat. But if it is green, then try and send the message again", vbCritical, "Not Connected"
    Exit Sub
End If

' Check to see if our message isn't an action
If Left(txtSend.Text, 7) <> "/action" And Left(txtSend.Text, 3) <> "/me" Then
        Broadcast "msgø+username+ø" & txtSend.Text & "ø" & MessageColour & "ø" & MessageBold & "ø" & MessageUnderline '& "ø" & MessageHColour
        ' Is this message the same as your last one?
        If OldData <> txtSend.Text Then
        ' No? Then give you an NCredits
        ' and remember your last message sent
            OldData = txtSend.Text
            NCredits = NCredits + 1
        End If
End If

' Is our message an /action or /me?

' Because of the different lengths of the string
' (ie. /action = 7, /me = 3), we need to have
' different actions for them, namely the mid part
If Left(txtSend.Text, 7) = "/action" Then
    Broadcast "actø+username+ø" & Mid(txtSend.Text, 8)
        If OldData <> txtSend.Text Then
            OldData = txtSend.Text
            ' Sending an /action or /me gives you 2!!
            NCredits = NCredits + 2
        End If
ElseIf Left(txtSend.Text, 3) = "/me" Then
    Broadcast "actø+username+ø" & Mid(txtSend.Text, 4)
        If OldData <> txtSend.Text Then
            OldData = txtSend.Text
            ' Sending an /action or /me gives you 2!!
            NCredits = NCredits + 2
        End If
End If

' Adds the last message to the chat 'history' box
txtSend.AddItem txtSend.Text, 0
txtSend.Text = ""
End Sub

Private Sub DoLock()
' When you set your status as AFK or Unavailable,
' then everything is locked off except for these
' objects. When you set your status as AFK
' or unavailable, we don't want you harassing
' the people in the room, do we? (Unless you
' are an administrator >:D )

' BTW, List1 is actually a ListView, not
' a generic list box. It is called List1
' because it saved changing all the associated code
List1.Enabled = True
mnuInfo.Enabled = True
mnuUnAvailable.Enabled = True
mnuAFK.Enabled = True
mnuAvailable.Enabled = True
mnuUE.Enabled = True
mnuDev.Enabled = True
mnuUEGhost.Enabled = True
mnuUEPIP.Enabled = True
mnuUERedirect.Enabled = True
mnuUEKick.Enabled = True
mnuUEAdmin.Enabled = True
mnuUERemAdmin.Enabled = True
mnuRawData.Enabled = True
mnuTray.Enabled = True
mnuTray_AFK.Enabled = True
mnuTray_Available.Enabled = True
mnuTray_Unavailable.Enabled = True
mnuTray_Show.Enabled = True
mnuTray_Quit.Enabled = True

End Sub



Private Sub Form_Load()
On Error Resume Next

' Tray is actually our class module, clsTray

' Initialize Syntax: Hwnd to create icon for, Icon
' to use, Default tooltip
Tray.Initialize Me.hwnd, Me.Icon, "NChat - Connecting..."
Tray.ShowIcon

' CaptionPrefix, is what the frmMain.caption starts
' as. Eg, if CaptionPrefix = "Hi-", then frmMain's
' caption would look like this: "Hi-Grayda has entered
' the room". If EndMessage is "!!" then it would
' appear like this: "Hi-Grayda has entered the room!!"
CaptionPrefix = "Welcome to NChat Alpha!! - "
'MkDir App.Path & "\Recieved Files"
' End window is the window's title suffix
EndWindow = "!!"
' Stops other programs from stealing your port :)
'sckUDP.Bind sckUDP.LocalPort, sckUDP.LocalIP
frmLog.Show
frmLog.Visible = False

' Clears the chat text and displays some headings
mnuClear_Click

' Command Line stuff. use -r#### to load a new room
' so you can administer 2+ rooms at the same time
If Left(Command$, 2) = "-r" Then
' NewRoom closes the current connection, changes the port
' and then reconnects, in a different "room".
' The second part of the command is the name of the room
' and the third part is whether or not to announce the
' room change (That is optional)
NewRoom Val(Mid(Command$, 3)), "Startup room"
Else
' If you aren't connecting to another room, then
' Tell the user they have 2 open
If App.PrevInstance = True Then
MsgBox "You already have a copy of NChat running!! Please close it down to stop conflict!", vbCritical, "NChat already open!"
End
End If
End If

NewRoom "4442", "Lobby", True

' Set up our file Receiver
Receiver.BinaryReceiver1.Listen

' Makes a directory called Profiles, and copies
' All the Profiles from the RES file into the folder
 MkDir AppPath & "Profiles"

For i = 101 To 114
If FileObj.FileExists(AppPath & "Profiles\" & LoadResString(i) & ".pro") = False Then CopyFromRes i, "PROFILES", LoadResString(i) & ".pro", "\Profiles\"

Next i

' Loads settings from your settings file
LoadSettings
Close #1

' No start / end msg thing? Set one!
' This only happens when you are running NChat
' for the first time, or your settings are corrupt
If StartMSG = "" Then StartMSG = "||"
If EndMSG = "" Then EndMSG = "||"

' The Bigger TotalIcons is, then
' the more user icons you can select from frmOptions
If TotalIcons <= 18 Then TotalIcons = 19

' MyIcon is the little icon that appears next
' to your username on the list to the right

' No MyIcon set? The set it as #1
If MyIcon = 0 Or MyIcon > ImageList1.ListImages.Count - 1 Then MyIcon = 1

' Tell the room you have connected.
' There was a problem recieving your own
' "con" data, so that's why this is texted
Broadcast "conø" & UserName & "ø" & Val(MyIcon) & "ø" & sckUDP.LocalIP
Text "+username+ has joined the chat room!" & vbCrLf, con, True, , , , vbCenter
Status CaptionPrefix & UserName & " has joined the chat room!" & EndWindow
' And adds your name to the top of the list
List1.ListItems.Add 1, frmMain.sckUDP.LocalIP, UserName, , ImageList1.ListImages.Item(MyIcon).Key

' Resize just to keep everything in shape
Form_Resize

' For more info on how to tell how fast a program is
' loaded, check modMisc, and look at the GetTickCount
' Public Declare.
Log "NChat sucessfully loaded in: " & GetTickCount - OldTickCount & " milliseconds" & vbCrLf & vbCrLf, vbBlack, True

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
' Code I swiped from somewhere (I think PSC :D)
' That lets you right click on a tray icon

' I don't understand it too well, but it works! :)
Dim msgCallBackMessage As Long
msgCallBackMessage = X / Screen.TwipsPerPixelX
Dim WM_RBUTTONDOWN As Long
Dim WM_RBUTTONUP As Long
Dim WM_LBUTTONDBLCLK As Long

WM_LBUTTONDBLCLK = &H203
WM_RBUTTONDOWN = &H204
WM_RBUTTONUP = &H205

Select Case msgCallBackMessage
    Case WM_RBUTTONUP
        Me.PopupMenu mnuTray
    Case WM_LBUTTONDBLCLK
        If Button <= 0 Then Exit Sub
        If mnuTray_Show.Visible = False Then
        mnuHide_Click
        Else
        mnuTray_Show_Click
        End If
        Case Else
        Exit Sub
    End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Form_Unload (0)
End Sub

Private Sub Form_Resize()
On Error Resume Next

' Just some resize stuff
txtChat.Move 150, 150, Me.Width - 400 - List1.Width - 150, Me.Height - 1400 - SB1.Height
txtSend.Move 150, txtChat.Height + 250, txtChat.Width
cmdSend.Move txtSend.Width + 400, txtSend.Top
cmdAction.Move cmdSend.Left + cmdSend.Width + 50, cmdSend.Top
List1.Move txtChat.Width + 300, txtChat.Top, List1.Width, txtChat.Height - 1100
ListView1.Move List1.Left, List1.Height + 250, List1.Width
lstIgnore.Move ListView1.Left, ListView1.Top, ListView1.Width

End Sub

Private Sub Form_Terminate()
Form_Unload (0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
' Get rid of the tray icon
Tray.HideIcon
' Disconnects you from the server
Broadcast "disø+username+"
' Opens your settings file (Or creates one)
' Then encodes and writes your settings
Open AppPath & GetUserName & ".ncg" For Output As #1
Print #1, CBool(ShareMyFiles)
Print #1, frmFileList.File1.Path
Print #1, frmFileList.File1.Pattern
' These are your "Information" settings
' I wanted to call them profiles, but
' we already have Profiles (Skins)
UserName = Replace(UserName, "[RA]", "")
Print #1, Encode(UserName, GetUserName & "©®©32")
' Total Icons is how many User Icons you have bought from
' the NChat Shop
Print #1, Encode(TotalIcons, GetUserName & "©®©32")

' Because the room admin box isn't opened
' when you restart NChat (Check out NewRoom. You
' will see that it does. There are also security
' reasons)
' Your Username
If frmMain.mnuDev.Visible = True Then
' About AdminAhoy and TharSheBlows: I was watching
' the Simpsons and Handsome Pete was on, dancing.
' Bart threw him a quarter and Pete continued dancing.
' The Sea Caption said: "NOT A QUARTER! He'll be dancin' for
' hours!" :D
Print #1, Encode("AdminAhoy", GetUserName & "©®©32")
ElseIf frmMain.mnuDev.Visible = False Then
Print #1, Encode("TharSheBlows", GetUserName & "©®©32")
End If

' How many NCredits your have
Print #1, Encode(NCredits, GetUserName & "©®©32")

' Your Windows Username. When your settings are loaded, this
' is checked against your current username. If it doesn't
' match, then you are kicked, and told you are stealing credits
Print #1, Encode(GetUserName, GetUserName & "©®©32")
' The icon that appears next to your name
Print #1, Trim(MyIcon)
' Whether or not you can see swearing that comes in
Print #1, Swearing
' Your last profile loaded
If IniFile(1) = "" Then IniFile(1) = "BLANK"
Print #1, Encode(IniFile(1), GetUserName & "©®©32")
' Are you a TRUE admin? True admins can't be kicked
Print #1, Encode(TrueAdmin, GetUserName & "©®©32")
' Do Private Messages popup?
Print #1, Popup
' Show the tip of the day?
Print #1, DontShowTip

' The fancy username features, such as boldness, underlining, colour and
' highlight colour
Print #1, Encode(MessageBold, GetUserName & "©®©32")
Print #1, Encode(MessageUnderline, GetUserName & "©®©32")
Print #1, Encode(MessageColour, GetUserName & "©®©32")
Print #1, Encode(MessageHColour, GetUserName & "©®©32")

' The Start of your message: StartMSG --> || YourName || <-- EndMSG Your message etc.
Print #1, StartMSG
Print #1, EndMSG
If Ban = True Then
Print #1, Encode("Banned", "12345678910")
End If
Close #1

' Fade the window out. Thanks to allapi.net for
' this code :). This also lets you know if
' your settings have been saved
AnimateWindow Me.hwnd, 200, AW_HIDE Or AW_BLEND
End
End Sub

Private Sub List1_DblClick()
' When you double click a name in the list of users, it opens up a new chat box
If List1.SelectedItem.Text > "" And List1.SelectedItem.Text <> UserName And List1.ListItems.Count > 0 Then ul1_Click
End Sub

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
'On Error Resume Next

' If it's not a left click, or it's blank, then ignore it
If List1.SelectedItem.Text = "" Or Button <> 2 Or List1.ListItems.Count = 0 Then Exit Sub 'Or List1.SelectedItem.Text = Username Then Exit Sub
If List1.SelectedItem.Text = UserName Then ' And mnuDev.Visible = False Then
' It's a menu, so it pops up under the cursor
' This menu has your status, ignore, change
' username and icon stuff.
PopupMenu mnuInfo

' Without this, when you right click on your username,
' The first menu pops up, then when you click away, the
' second one pops up, posing a potential security risk
Exit Sub
End If

' AFK is Away From Keyboard.
' If they are afk, then you can't send 'em messages
' unless you are an admin
If mnuDev.Visible = False And List1.SelectedItem.SmallIcon = ImageList1.ListImages.Item(ImageList1.ListImages.Count).Key Then
Text "This user is AFK, and cannot be contacted" & vbCrLf, ThatsBad, True
Exit Sub
End If
If mnuDev.Visible = False And List1.SelectedItem.SmallIcon = ImageList1.ListImages.Item(ImageList1.ListImages.Count - 1).Key Then
Text "This user is unavailabe, and cannot be contacted" & vbCrLf, ThatsBad, True
Exit Sub
End If
' Replaces >> with the username, so you know who
' you clicked on
If List1.SelectedItem.Index > -1 And List1.SelectedItem.Text > "" Then
ul1.Caption = "Send " & List1.SelectedItem.Text & " a Private Message"
ul2.Caption = "Send " & List1.SelectedItem.Text & " some NCredits"
ul4.Caption = "Browse " & List1.SelectedItem.Text & "'s Files"
'ul5.Caption = "Add " & List1.SelectedItem.Text & " to friend list"
mnuUEIgnore.Caption = "Ignore " & List1.SelectedItem.Text

' Are you an admin? if you are, then show the 'extra' menu
If mnuDev.Visible = True Then mnuUE.Visible = True
PopupMenu mnuUserList
End If

End Sub

Private Sub ListView1_DblClick()
' When we double click our ListView1 (Which is
' our private message box), NChat detects
' which username you clicked on, and then
' opens private messages ONLY from them, leaving
' the others there.

' Temporary Array for private messages
Dim Splice() As String
Dim b As Integer

On Error GoTo Done
' Not a message? Don't continue
If ListView1.SelectedItem.Text = "" Then Exit Sub
' Load our data into our array and SplitVB5 it
Splice = SplitVB5(ListView1.SelectedItem.Key, "…")
For i = LBound(CW) To UBound(CW)

' Find an unused window
If CW(i).Tag = "" Then
CW(i).Show
' Add text etc.
Txt2 ListView1.SelectedItem.Text & " ::  " & Splice(1) & vbCrLf, act, Int(i)
CheckRTF CW(i).Text1

' The tag allows us to see who we are chatting to
CW(i).Tag = ListView1.SelectedItem.Text

' Clear the list for the next person
ListView1.ListItems.Remove (List1.SelectedItem.Index)

If ListView1.ListItems.Count = 0 Then Exit Sub
b = 0
For n = LBound(CW) To UBound(CW)
Do Until b = ListView1.ListItems.Count
b = b + 1
If ListView1.ListItems.Item(b).Text = CW(n).Tag Then
ListView1.ListItems.Item(b).Selected = True
ListView1_DblClick
b = 0
End If
Loop
Exit Sub
Next n

ElseIf CW(i).Tag = ListView1.SelectedItem.Text Then

Txt2 ListView1.SelectedItem.Text & " ::  " & Splice(1) & vbCrLf, act, Int(i)
CheckRTF CW(i).Text1
ListView1.ListItems.Remove (List1.SelectedItem.Index)
For b = 1 To ListView1.ListItems.Count
If ListView1.ListItems.Item(i).Text = CW(i).Tag Then
ListView1.ListItems.Item(i).Selected = True
ListView1_DblClick
End If
Next b

Exit Sub
End If
Next i
Done:
If Err.Number = 35600 Then Exit Sub
End Sub

Public Sub LoadProfile()
'On Error Resume Next
' The loading of the profile
' Put into one easy to use sub,
' so the code doesn't have to be repeated

EndWindow = ""
'tsk. so many doevents... sad.
DoEvents
DoEvents

' See modPublic for descriptions of these
Heading = ReadText("Theme", "Heading", 1)
Msg = ReadText("Theme", "Message", 1)
dis = ReadText("Theme", "Disconnect", 1)
svr = ReadText("Theme", "Server", 1)
act = ReadText("Theme", "Action", 1)
con = ReadText("Theme", "Connect", 1)
ThatsGood = ReadText("Theme", "Good", 1)
ThatsBad = ReadText("Theme", "Error", 1)
frmMain.BackColor = ReadText("Theme", "Background", 1)
frmMain.txtChat.BackColor = ReadText("Theme", "ChatBackColor", 1)
frmMain.txtSend.BackColor = ReadText("Theme", "SendBackColor", 1)
frmMain.txtSend.ForeColor = ReadText("Theme", "SendForeColor", 1)
DoEvents

Tmp = ReadText("Theme", "Title", 1) & " "
If Trim(Tmp) > "" Then CaptionPrefix = Tmp
EndWindow = " " & ReadText("Theme", "EndWindow", 1)
Status CaptionPrefix & " New Profile Loaded! " & EndWindow
mnuClear_Click
' The picture is loaded last, incase there is a problem
If ReadText("Theme", "BackgroundPic", 1) > "" Then picRed.Picture = LoadPicture(ReadText("Theme", "BackgroundPic", 1))

End Sub

Private Sub LoadSettings()
' Loads your NCredits, Username etc.
' Most things need to be decoded before they are loaded
On Error Resume Next
Dim TempStr As String

' Pretty obvious: Opens your file and reads from it
Open AppPath & GetUserName & ".ncg" For Input As #1
' Whether or not your files are shared
Line Input #1, TempStr
ShareMyFiles = CBool(TempStr)
' What folder you have shared
Line Input #1, TempStr2
frmFileList.File1.Path = TempStr2
Line Input #1, TempStr
' What files you have shared (Like *.mp3;*.jpg etc.)
FilePattern = TempStr
' Some simple user info stuff. Read the variables
' to find out what is loaded
Line Input #1, TempStr
UserName = Decode(TempStr, GetUserName & "©®©32")
' No username set? Then it's set as your windows username
If Trim(UserName) = "" Then
UserName = GetUserName
' using windows 95 / 98 or no windows username set? Then
' why not use our IP address? :D
If Trim(UserName) = "" Then UserName = sckUDP.LocalIP
' Uh, if you have no username, then you are a new user
' so show the tip of the day box
frmTip.Show
frmMain.Enabled = True

End If

' Number of user icons you have "Purchased" from the NChat
' shop (THere are like 50 of them to buy!!)
Line Input #1, TempStr
TotalIcons = Decode(TempStr, GetUserName & "©®©32")

' Whether or not you are an admin
Line Input #1, TempStr
If Decode(TempStr, GetUserName & "©®©32") = "AdminAhoy" Then
mnuDev.Visible = True
Else
mnuDev.Visible = False
End If

' How many NCredits you have saved
' (Think of NCredits like $$. Save em to buy things)
Line Input #1, TempStr
NCredits = Decode(TempStr, GetUserName & "©®©32")

DoEvents
' Username Check Description:

' Username is your NChat username
' UsernameCheck is your windows username

' If usernamecheck is different than your windows
' username, then one of two things have happened:

' 1) You have changed your windows username
' 2) You took someone else's file and called it yours

' When this happens, you get humiliated and stripped of NCredits
Line Input #1, TempStr
UsernameCheck = Decode(TempStr, GetUserName & "©®©32")
DoEvents

If UsernameCheck = "" Then Exit Sub
If UsernameCheck <> GetUserName Then
Kill AppPath & UserName & ".ncg"
NCredits = 0
frmCheating.Show
Do Until frmCheating.Visible = False
DoEvents
DoEvents
Loop
mnuDev.Visible = False
TrueAdmin = False

' humiliation!!
Broadcast ("svrø+username+ tried to earn NCredits dishonestly...")
Exit Sub
'End ' Uncomment this line to kick them instead
End If

Line Input #1, TempStr
' Your icon next to your username
MyIcon = TempStr
Line Input #1, TempStr
' Is swearing filtered out?
Swearing = CBool(TempStr)
Line Input #1, TempStr
'Line Input #1, TempStr
' Your last Profile (INI) Loaded
IniFile(1) = Decode(TempStr, GetUserName & "©®©32")
If IniFile(1) > "" And IniFile(1) <> "True" And IniFile(1) <> "BLANK" Then LoadProfile
Line Input #1, TempStr
' See top of this form code for trueadmin info
TrueAdmin = Decode(TempStr, GetUserName & "©®©32")
Line Input #1, TempStr
' Popup Private Messages (Rather than hide them)?
Popup = CBool(TempStr)

Line Input #1, TempStr
' Show the tip of the day?
If TempStr = "False" Or TempStr = "" Then
frmTip.Show
OnTop frmTip.hwnd
Else
DontShowTip = True
End If

' TT1 is assigned, because for some reason, the
' code called a byval error for a dimmed integer,
' which was strange. This fixes the problem for now
TT1 = GetUserName
Line Input #1, TempStr
' When you send a message, is your username BOLD?
MessageBold = Decode(TempStr, TT1 & "©®©32")
Line Input #1, TempStr
' When you send a message, is your username -UNDERLINED-?
MessageUnderline = Decode(TempStr, TT1 & "©®©32")
Line Input #1, TempStr
' When you send a message, is your username a different colour?
MessageColour = Decode(TempStr, TT1 & "©®©32")
Line Input #1, TempStr
' When you send a message, is your username highlighted in a cool colour?
MessageHColour = Decode(TempStr, TT1 & "©®©32")
Line Input #1, TempStr
' The start of the message (eg. || Username || Hello!)
StartMSG = TempStr
Line Input #1, TempStr
' The end of the message, like Startmsg
EndMSG = TempStr

Line Input #1, TempStr
' Are you permanently banned from NChat?
Tmp = Decode(TempStr, "12345678910")
If Tmp = "Banned" Then
MsgBox "YOU HAVE BEEN BANNED FROM NCHAT. NCHAT WILL NOW CLOSE", vbCritical, "BANNED"
End
End If
Close #1

End Sub

Private Sub lstIgnore_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error Resume Next
' Popup the ignore menu, where you can delete ignored people on your list
If Button = 2 Then PopupMenu mnuPopup_Ignore

End Sub

Private Sub mnuAbout_Click()
' New in this build: New About Box. Used to use API to show
' windows about box, but scrapped that idea, coz I needed to thank
' some people etc.
frmAbout.Show
End Sub

Private Sub mnuAdminWindow_Click()
' Show the sData log window. Lets you view all incoming NChat data
frmLog.Show
End Sub

Private Sub mnuAFK_Click()
' Locks NChat down so you can't be afk and send messages
On Error Resume Next
Broadcast "chiø" & UserName & "ø" & ImageList1.ListImages.Count
Dim Obj As Object
For Each Obj In frmMain
Obj.Enabled = False
Next

DoLock
End Sub

Private Sub mnuAutoBot_Click()
' Starts / Stops / Shows the NChat Automatic Robot, Called the AutoBot, or
' Notch, as is he / she / it is known
frmAutobot.Show
End Sub

Private Sub mnuAvailable_Click()
' Sets your status to available
On Error Resume Next
' Chi Lets everyone know that your Icon has changed.
' change icon syntax: chi <delimiter>, username <delimiter> new icon
Broadcast "chiø" & UserName & "ø" & MyIcon
' UnLocks NChat
Dim Obj As Object
For Each Obj In frmMain
Obj.Enabled = True
Next
DoLock
txtSend.Text = ""
End Sub

Private Sub mnuAwayMSG_Click()
' Away message is like leaving a note for any potential
' Private Messengers. It is also like an answering machine
If AwayMessage = False Then

AwayMSG = InputBox("Enter an Away Message. If someone tries to contact you in a private chat, this message will be displayed", "Away Message", AwayMSG)
mnuAwayMSG.Checked = True
AwayMessage = True
Else
mnuAwayMSG.Checked = False
AwayMessage = False
End If
End Sub

Private Sub mnuBalance_Click()
' < 0 NCredits, red text > 0 NCredits green text
If NCredits <= 0 Then
Text "You have +ncredits+ NCredits remaining" & vbCrLf, ThatsBad, True
Else
Text "You have +ncredits+ NCredits remaining" & vbCrLf, ThatsGood, True
End If
End Sub

Private Sub mnuBL_Click()
' Broadcast Mode lets you chat to EVERYONE on the network. Loopback lets
' NChat connect to itself, making it useful for testing on one computer,
' if you don't have a valid network
If MsgBox("WARNING: BY SWITCHING TO LOOPBACK MODE, NCHAT CAN CONNECT TO ITSELF, AND NOT WITH OTHER NCHATS ON OTHER COMPUTERS. ONLY DO THIS IF YOU KNOW WHAT YOU ARE DOING. IF YOU ARE UNSURE, CLICK CANCEL", vbCritical + vbYesNo, "WARNING: READ CAREFULLY") = vbNo Then Exit Sub
If Loopback = False Then
Address = "127.0.0.1"
Loopback = True
MsgBox "Now in loopback mode. To change back, click button again", vbInformation, "Loopback"
frmMain.SB1.Panels(1).Text = "NChat - Online!!"
frmMain.SB1.Panels(3).Picture = frmMain.picGreen.Picture
Else

Address = "255.255.255.255"
Loopback = False
MsgBox "Now in broadcast mode. To change back, click button again", vbInformation, "Loopback"
frmMain.SB1.Panels(1).Text = "NChat - Online!!"
frmMain.SB1.Panels(3).Picture = frmMain.picGreen.Picture

End If
End Sub

Private Sub mnuChangeNCredits_Click()
' Change your NCredits, and stops any overflow errors
On Error Resume Next
NCredits = InputBox("How many NCredits would you like?", "Change NCredits", NCredits)
End Sub

Private Sub mnuChangeRoomInfo_Click()
' Is as it says. Changes your room info if you make a spelling mistake
Tmp = InputBox("Please enter the new name for your room", "New Room Name", RoomName)
If Trim(Tmp) = "" Then
Text "Room name was blank. Not Changed" & vbCrLf, ThatsBad, True
Exit Sub
ElseIf Trim(Tmp) = RoomName Then
Text "Room name is the same as your old one. Not Changed" & vbCrLf, ThatsBad, True
Exit Sub
End If
Text "This room (" & RoomName & ") has been changed to " & Tmp & "." & vbCrLf, ThatsGood, True
Broadcast "svrøThis room's name is now set as: " & Tmp & " (Was " & RoomName & ")"
RoomName = Tmp

End Sub

Private Sub mnuClear_Click()
txtChat.Text = ""
' Some fancy headings, has no purpose of course,
' just looks good :)
Text "Welcome to NChat ", Heading, True, False, False, 18, vbCenter, , "False"
Text "Alpha!!" & vbCrLf, ThatsGood, True, False, False, 18, vbCenter, , "False"
Text "Created by Grayda of Solid Inc." & vbCrLf, Heading, , , , 7, vbCenter, , "False"
Text ":nchat" & vbCrLf & "http://www.solidinc.tk" & vbCrLf, Heading, False, False, False, 7, vbCenter, , "True"
Text "Preparting to Start NChat Alpha, Please stand by..." & vbCrLf, Heading, True, False, False, 8, vbCenter, , "True"

' IniFile(x) allows you to have more than one INI File
' open at a time
If IniFile(1) = "" Then Exit Sub
'LoadProfile
Tmp = ReadText("Theme", "Description", 1)
Text Tmp & vbCrLf, Heading, True, , , , vbCenter
' Profile Welcome Message
For i = 0 To 50
Tmp = ReadText("Theme", "WelcomeMSG" & i, 1)
' Blank? Don't text it
If Tmp > "" Then Text Tmp & vbCrLf, Heading, True, , True, 8, vbCenter
Next i

End Sub

Private Sub mnuConnect_Click()
frmServerConnect.Show

End Sub

Private Sub mnuCreateRoom_Click()
frmCRI.Show
End Sub

Private Sub mnuDeleteIgnore_Click()
If lstIgnore.SelectedItem.Text > "" Then lstIgnore.ListItems.Remove (lstIgnore.SelectedItem.Text)
End Sub

Private Sub mnuDisconnect_Click()
IsInternet = False
Address = "255.255.255.255"
sckUDP.Close
sckRooms.Close
sckUDP.Protocol = sckUDPProtocol
sckRooms.Protocol = sckUDPProtocol
sckUDP.Connect
sckRooms.Connect
mnuChatRooms.Visible = True

End Sub

Private Sub mnuDoHeading_Click()
' Large 'message' (Heading)
Broadcast "heaø" & InputBox("What do you want your heading to say?", "Heading") & "ø" & UserName
' 'Resets' the text so the next message isn't BIG
Text "" & vbCrLf, , , , , 8

End Sub


Private Sub mnuHide_Click()
mnuTray_Show.Visible = True
mnuHide.Visible = False

Me.Visible = False
Tray.Box "NChat has been minimized to the tray! To restore it, double click the NChat icon!", "NChat"
End Sub

Private Sub mnuIgnore_Click()
mnuIgnore2_Click

End Sub

Private Sub mnuIgnore2_Click()
ListView1.Visible = False
lstIgnore.Visible = True



mnuIgnore2.Checked = True
mnuPM.Checked = False


End Sub

Private Sub mnuInsertFake_Click()
' User (Fake or real) connect syntax:
'  con <Delmiter> Username <Delmiter> Icon #
On Error Resume Next
Dim Icon As Integer
Randomize
FakeUser = InputBox("Enter fake user's name", "Fake user")
Icon = InputBox("Enter the INDEX number of their icon: 1-" & ImageList1.ListImages.Count, "Icon?", Int(Rnd * ImageList1.ListImages.Count))
If FakeUser > "" And Icon >= 1 Then
Broadcast "conø" & FakeUser & "ø" & Icon & "ø" & Int(Rnd * 255) & "." & Int(Rnd * 255) & "." & Int(Rnd * 255) & "." & Int(Rnd * 255) & "." & Int(Rnd * 255)
'Tex6t FakeUser & " has entered the conversation!" & vbCrLf, con, True, , , , vbCenter
'List1.ListItems.Add 2, FakeUser, FakeUser, , ImageList1.ListImages.Item(Icon).Key
NewUser = FakeUser
End If
End Sub

Private Sub mnuJoinCustom_Click()
On Error Resume Next
frmRooms.List1.ImageList = ImageList1

frmRooms.List1.Nodes.Clear
frmRooms.List1.Nodes.Add , , "N", "NChat Rooms", frmMain.ImageList1.ListImages.Item(18).Key
frmRooms.List1.Nodes.Add , , "C", "User-Made Rooms", frmMain.ImageList1.ListImages.Item(18).Key
frmRooms.List1.Nodes.Item(1).Expanded = True
frmRooms.List1.Nodes.Item(2).Expanded = True
RI = frmMain.ImageList1.ListImages.Item(7).Key
frmRooms.List1.Nodes.Add "N", 4, "R4442", "Lobby", RI
frmRooms.List1.Nodes.Add "N", 4, "R4443", "Music Chat", RI
frmRooms.List1.Nodes.Add "N", 4, "R4444", "The Work Room", RI
frmRooms.List1.Nodes.Add "N", 4, "R4445", "Help for NChat", RI
frmRooms.List1.Nodes.Add "N", 4, "R4446", "Programmers Chat", RI
frmRooms.List1.Nodes.Add "N", 4, "R4447", "Room for fighters", RI
RoomBroadcast "lst"
frmRooms.Show
End Sub

Private Sub mnuLoadProfile_Click()
' Loads profiles (skins)
On Error Resume Next
OldINI = IniFile(1)
dlgSave.InitDir = AppPath & "Profiles\"
dlgSave.Filter = "INI Files (*.pro)|*.pro"
dlgSave.DialogTitle = "Select Profile to load..."

dlgSave.ShowOpen
If dlgSave.FileTitle = "" Then Exit Sub
IniFile(1) = dlgSave.filename

' See one of the modules for help with loadprofile
LoadProfile

' Is your old INI the same as your new one?
' No 10 NCredits for you!
If OldINI <> IniFile(1) Then NCredits = NCredits + 10
Text "You have recieved 10 NCredits for loading a profile!!" & vbCrLf, ThatsGood, True, , , , vbCenter
End Sub

Private Sub mnuNChat_Stats_Click()

End Sub

Private Sub mnuNewProfile_Click()
frmNewProfile.Show
End Sub

Private Sub mnuNewRoom_Click()
' Creates a room. Doesn't have a random # like above
Tmp = InputBox("Enter the room to create. You can choose the room number", "Custom Room")
RoomName = InputBox("Enter your room's name. You can choose the new name", "Custom Room")
If Tmp > "" Then
CreatedRoom = True

NewRoom Tmp, RoomName
frmAdmin.Show
End If

End Sub

Private Sub mnuOptions_Click()
frmOptions.Show

End Sub

Private Sub mnuPM_Click()

ListView1.Visible = True
lstIgnore.Visible = False

mnuIgnore2.Checked = False
mnuPM.Checked = True


End Sub

Private Sub mnuQuit_Click()
Form_Unload (0)
End Sub

Private Sub mnuRawData_Click()
Dim tmp2 As String

' Raw data refers to data that doesn't have a prefix such
' as hea, msg, cdo ect.
tmp2 = InputBox("Enter raw data to broadcast. MUST include delimiters for standard messages (Alt+0248 or +d+)...", "Raw Data")
If tmp2 > "" Then Broadcast tmp2
End Sub

Private Sub mnuRealUser_Click()
' isr = Is Real? Checks if a user is real or fake
Broadcast "isr+d+" & List1.SelectedItem.Text
End Sub

Private Sub mnuRemFake_Click()
' Removes a fake user from the room
Broadcast "disø" & InputBox("Enter username to disconnect", "Remove Fake User", NewUser)
End Sub

Private Sub mnuSaveChat_Click()
' Saves chat data
' If the selected type is *.txt or *.*
' then save it as plain text.
' If it's not, then save it with RTF tags etc.
dlgSave.Filter = "Text Files (*.TXT)|*.txt|RTF File (*.rtf)|*.RTF|All Files (*.*)|*.*"
dlgSave.ShowSave
If dlgSave.filename = "" Then Exit Sub
Open dlgSave.filename For Output As #1
If dlgSave.FilterIndex = 1 Or dlgSave.FilterIndex = 3 Then
Print #1, txtChat.Text
ElseIf dlgSave.FilterIndex = 2 Then
Print #1, txtChat.TextRTF
End If
Close #1
Text "File Saved!" & vbCrLf, svr, True
End Sub

Private Sub mnuSendAll_Click()
Tmp = InputBox("Enter text to send. This will be sent to EVERYONE using NChat, even people in other rooms", "Send to ALL Rooms")
If Tmp > "" Then sckRooms.SendData "comø" & Tmp
End Sub


Private Sub mnuSendPic_Click()
'On Error Resume Next
If MsgBox("Warning: This method of picture transfer is VERY unstable. MOST pictures will NOT be sent, and will take more than 4 minutes to transfer ANY picture OVER 30kb. Use at your own risk! Do you wish to continue?", vbCritical + vbYesNo, "Send Picture?") = vbNo Then Exit Sub

Dim t As StdPicture
dlgSave.ShowOpen


SendPic LoadPicture(dlgSave.filename, , 2)

End Sub

Private Sub mnuSmileys_Click()
' Sometimes all the smileys won't show up, so do it twice
' to ensure it shows up ok
frmSmileys.Show
Unload frmSmileys
frmSmileys.Show
End Sub



Private Sub mnuStore_Click()
frmStore.Show
End Sub

Private Sub mnuTextHelp_Click()
If FileObj.FileExists(AppPath & "Help\NChat Help.hlp") = True Then
ShellExecute 0&, "Open", AppPath & "Help\NChat Help.hlp", "", vbNullString, SW_SHOWNORMAL
Else
MsgBox "Cannot find " & AppPath & "Help\NChat Help.hlp. The help file cannot be opened. You can download it off the web-site at: http://www.solidinc.tk, under the Downloads section", vbCritical, "Help File Not Found!!"
End If


End Sub

Private Sub mnuTipofTheDay_Click()
frmTip.Show

End Sub

Private Sub mnuTray_AFK_Click()
mnuAFK_Click
End Sub

Private Sub mnuTray_Available_Click()
mnuAvailable_Click
End Sub

Private Sub mnuTray_Quit_Click()
Form_Unload (0)
End Sub

Private Sub mnuTray_Show_Click()
Me.Visible = True
Me.WindowState = 0
mnuHide.Visible = True
mnuTray_Show.Visible = False
End Sub

Private Sub mnuTray_Unavailable_Click()
mnuUnAvailable_Click
End Sub

Private Sub mnuUEAdmin_Click()
' Add admin (This menu is called from List1 on right click

' UE = User... um... Environment?... Education? Never mind :)
Broadcast "addø" & List1.SelectedItem.Text & "ø" & UserName
End Sub

Private Sub mnuUECHI_Click()
' Change your icon without going to options etc.
Tmp = InputBox("Enter a number between 1 and " & TotalIcons & " as your Icon", "Change User Icon", MyIcon)
If Tmp = MyIcon Then
Text "Your Icon is the same as your old one!" & vbCrLf, ThatsBad, True
Exit Sub
End If

' Cannot enter a -tive number
If Tmp < 1 Or Tmp = "" Then
Text "Your New Icon is less than 1!" & vbCrLf, ThatsBad, True
Exit Sub
End If

' Only admins can access all icons
If Tmp > ImageList1.ListImages.Count Then
Text "Your New Icon is more than " & ImageList1.ListImages.Count & "!" & vbCrLf, ThatsBad, True
Exit Sub
End If

' Cannot enter a # more than the number of icons you own
If Tmp > TotalIcons And mnuDev.Visible = False Then
Text "Your New Icon is more than " & TotalIcons & "!" & vbCrLf, ThatsBad, True
Exit Sub
End If


MyIcon = Tmp
' Tell the room of your icon change
Broadcast "chi+d+" & UserName & "+d+" & MyIcon
End Sub

Private Sub mnuUECHU_Click()
' CHange username
' Admins can have [A] and [RA] on their name
' but not wannabes
OldUsername = UserName
Tmp = InputBox("Enter new username. It has to be different to your old one, and cannot be blank", "Change Username", UserName)
If Tmp = "" Then
Text "Your New Username is blank!!" & vbCrLf, ThatsBad, True
Exit Sub
End If

If Right(Trim(Tmp), 3) = "[A]" And frmMain.mnuDev.Visible = False Or Right(Trim(Tmp), 4) = "[RA]" And frmMain.mnuDev.Visible = False Then
MsgBox "Sorry, but only administrators can have [A] or [RA] on the end of their name...", vbExclamation, "Bad Username"
Exit Sub
End If


If Tmp = UserName Then
Text "Your New Username is the same as your old one!" & vbCrLf, ThatsBad, True
Exit Sub
End If

UserName = Tmp
Broadcast "chuø" & OldUsername & "ø" & UserName & "ø" & MyIcon
DoEvents

Broadcast "svrø" & OldUsername & " is now known as " & UserName
Text "Changed your username!" & vbCrLf, Heading, True, , , , vbCenter


End Sub

Private Sub mnuUEGhost_Click()
' Ghosts (Message with someone else's username) a user
Broadcast "fakø" & List1.SelectedItem.Text & "ø" & InputBox("Enter message", "Ghost") & "ø" & UserName
End Sub

Private Sub mnuUEIgnore_Click()
' QuickIgnore(TM) :)
If MsgBox("Are you sure you want to ignore: " & List1.SelectedItem.Text & "? To unignore them, right click your name and select ignore list.", vbExclamation + vbYesNo, "Ignore User?") = vbYes Then lstIgnore.ListItems.Add , List1.SelectedItem.Text, List1.SelectedItem.Text, , ImageList1.ListImages.Item(ImageList1.ListImages.Count).Key

End Sub

Private Sub mnuUEKick_Click()
' KSV = Kick Server
' Kicks someone with a message from the admin
' (i.e The admin has kicked you etc.)
Broadcast "ksvø" & List1.SelectedItem.Text & "ø" & UserName
End Sub

Private Sub mnuUEPIP_Click()
' Print detailed info about a user
Broadcast "pipø" & List1.SelectedItem.Text & "ø" & UserName


End Sub

Private Sub mnuUERedirect_Click()
' Redirect a user.
RTMP = InputBox("Enter room number to redirect to:", "Redirect")
RTMP2 = InputBox("Enter room name", "Redirect", "Holiday Destination")
If RTMP > "" And RTMP2 > "" Then Broadcast "redø" & RTMP & "ø" & List1.SelectedItem.Text & "ø" & RTMP2 & "ø" & UserName
End Sub

Private Sub mnuUERemAdmin_Click()
' Kill an admin's rights
Broadcast "remø" & List1.SelectedItem.Text & "ø" & UserName
End Sub

Private Sub mnuUnAvailable_Click()
On Error Resume Next
' Makes sets your icon as 'unavailable'.
' Users cannot contact you while you are away
Broadcast "chiø" & UserName & "ø" & ImageList1.ListImages.Count - 1

' Locks NChat down so you can't be unavailable and send messages
Dim Obj As Object
For Each Obj In frmMain
Obj.Enabled = False
Next
AwayMessage = True
DoLock
End Sub

Private Sub mnuWmsg_Click()
' The welcome message is displayed on a user connect
' it doesn't have to be a message, but can be anything
' even a user limiter (i.e ksvø+newuser+)

' This is the same as DoStuff on entry, but
' only one command is sent
If WelcomeMsg = "" Then
WelcomeMsg = InputBox("Please enter new AutoWelcome Message", "WelcomeMsg", WelcomeMsg)
Broadcast ("svrøWelcome message set!!")
Else
WelcomeMsg = InputBox("Please enter updated AutoWelcome Message", "WelcomeMsg", WelcomeMsg)
Broadcast ("svrøWelcome message updated!!")
End If
End Sub






Private Sub sckRooms_DataArrival(ByVal bytesTotal As Long)

Dim RoomResult() As String
On Error Resume Next
Dim sckData As String
sckRooms.GetData sckData



RoomResult = SplitVB5(sckData, "ø")
frmLog.List1.AddItem sckData
Select Case RoomResult(0)

Case "lst"
If CreatedRoom = True Then
RoomBroadcast "roomøR" & sckUDP.LocalPort & "ø" & RoomName & "ø" & Int(MyIcon)
End If

Case "room"

For i = 1 To frmRooms.List1.Nodes.Count
If frmRooms.List1.Nodes.Item(i).Text = RoomResult(2) Then Exit Sub
Next i

If frmRooms.Visible = True Then frmRooms.List1.Nodes.Add "C", 4, RoomResult(1), RoomResult(2), ImageList1.ListImages.Item(Val(RoomResult(3))).Key


Case "info"
If RoomResult(1) = RoomName And CreatedRoom = True Then
RoomBroadcast "rriø" & RoomResult(2) & "øRoom Name: " & RoomName & "|||| " & Description & "||||Host: " & UserName & "||Using Version: " & App.Major & "." & App.Minor & "." & App.Revision & "||# of people in room: " & frmMain.List1.ListItems.Count & "||Room has been running for: " & RoomTime & " seconds||RoomCreate Software v: 2.1.0||Last person to enter: " & NewUser & "||NChat Room ID: " & sckUDP.RemotePort & "||Last message from: " & LastFrom
Text RoomResult(1) & " has requested your room info!" & vbCrLf, svr, True
End If

Case "rri"
If RoomResult(1) = UserName Then MsgBox Replace(RoomResult(2), "||", vbCrLf), vbInformation, "Room Information"

Case "com"
Text RoomResult(1) & vbCrLf, svr, True

Case Else
Exit Sub
End Select
End Sub



Private Sub sckUDP_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next

' Note: Please do not ask about TempData. TempData
' is so, because for some reason, sData does not traverse
' onto all forms, but tempdata does
sckUDP.GetData sData


' Once the first message has arrived, set the text again
' so no errors appear
SB1.Panels(1).Text = "NChat - Online!!"
SB1.Panels(3).Picture = picGreen.Picture
tempdata = sData

' Excuse the language. If swearing is off, then replace
' the language with more appropriate words
If Swearing = False Then
sData = Replace(sData, "shit", "sh*t")
sData = Replace(sData, "fuck", "f**k")
sData = Replace(sData, "cunt", "c*nt")
sData = Replace(sData, "bitch", "bi*ch")
End If

' A while ago, one person discovered that
' you could crash just about ANY chat program that used:
' 1) An updated version of RichEd32.dll
' 2) A rich text box control (Any sort seemed to work)
' just by sending a combination of text (usually starting
' with '{\RTF'). The trick was to send an 'object' (for example
' an OLE control) code. When it was discovered that the 'object'
' code didn't exist on the remote system, then the control would crash,
' along with the program it was attached to. You could fix the problem
' by using an old version of RichEd32.dll, but that was 100kb big, and
' wasn't handy for shipping with programs, so a smaller solution was to
' remove the {\ from the code, before the text reached the box and crashed
sData = Replace(sData, "{\", "##")
OldData = sData

' No Blank packets of data thanks!

Result = SplitVB5(sData, "ø")
If Trim(sData) = "" Or Result(0) <> "N©H@-|-" Then Exit Sub

sData = Replace(sData, "+newline+", vbCrLf)

' D is short for delimiter. Less keystrokes than pressin
' alt+0248
sData = Replace(sData, "+d+", "ø")

DoEvents
' Before NChat v4.0 (about March 2004), I didn't know
' what the Split command did! I have been programming
' for about 6-7 years, and only learnt in 2004
' what it does! Before that, you could only
' have commands that consisted of two parts.
' The first part MUST be 3 letters long, and the
' second part could be any amount of letters long.
' Except in exceptional circumstances, you couldn't
' have a third or fourth part :|

Select Case Result(1) ' Result(1) can be any length now
                      ' but 3 letters seems sufficient

' Result(2) is the username and Result(3) is the message
' Rather than sending || Username || Hello as the message,
' The ||'s are added on arrival
Case "msg"
' Remember: StartMSG = "||" until changed

' OK this section has been updated. If you
' purchase the items from the store, then you can
' have a bold, underlined, or even a different
' coloured text for your username

' This ensures Notch DOESN'T Talk until
' noone has talked for a certain amount of time
If frmAutoBotOptions.Timer1.Enabled = True Then
frmAutoBotOptions.Timer1.Enabled = Not frmAutoBotOptions.Timer1.Enabled
frmAutoBotOptions.Timer1.Enabled = Not frmAutoBotOptions.Timer1.Enabled
End If

Text StartMSG & " ", Msg
StartHere = Len(txtChat.Text)

' Text our username FIRST, then apply formatting later,
' so if there are errors, then they can be ignored
Text Result(2), Msg 'Int(Result(4)), CBool(Result(5)), False, CBool(Result(6))

Text " " & EndMSG & " ", Msg
Text Result(3) & vbCrLf, Msg

' Now for Text formatting! wooooo!
txtChat.SelStart = StartHere
txtChat.SelLength = Len(Result(2))
' Text highlighting has been removed. Caused some errors on Windows ME
'HighlightText (Result(7))
txtChat.SelColor = Int(Result(4))
txtChat.SelBold = CBool(Result(5))
txtChat.SelUnderline = CBool(Result(6))


LastFrom = Result(2)

' Log Logs the data into the frmLog textbox
Log "Message: " & Result(2) & " - " & Result(3) & vbCrLf, vbBlue
' Changes the caption to say the new message
Status CaptionPrefix & Result(2) & " - " & Result(3) & EndWindow

Case "act" ' action
If frmAutoBotOptions.Timer1.Enabled = True Then
frmAutoBotOptions.Timer1.Enabled = False
frmAutoBotOptions.Timer1.Enabled = True
End If

Text Result(2) & Result(3) & vbCrLf, act, True
Status CaptionPrefix & "A: " & Result(2) & Result(3) & EndWindow
Log "Action: " & Result(2) & Result(3) & vbCrLf, 33023

Case "svr" ' Server (Purple) message
Text Result(2) & vbCrLf, svr, True
Status CaptionPrefix & "S: " & Result(2) & EndWindow
Log "Server: " & Result(2) & vbCrLf, vbMagenta

Case "add" ' Add administrator
Log "Add Admin: " & Result(2) & " added by: " & Result(3) & vbCrLf, 32768

If Result(2) = UserName Then
mnuDev.Visible = True
Text Result(3) & " has made you an administrator!!" & vbCrLf, svr, True
Text "These powers have been given to you because " & Result(3) & " trusts you with them" & vbCrLf, svr, True
Text "Abusing these powers can see you get kicked, or even banned from NChat!" & vbCrLf, svr, True
OldUsername = UserName
UserName = Replace(UserName, " [A]", "")
UserName = UserName & " [A]"
Broadcast "chuø" & OldUsername & "ø" & UserName
End If

Case "startpic"
Text "" & vbCrLf
ThePic = ""
PicStarted = True

Case "endpic"
PicStarted = False
ThePic = Trim(Replace(ThePic, "##", "{\"))
If Right(Trim(ThePic), 2) <> "}}" Then ThePic = ThePic & "}}"
txtChat.SelStart = Len(txtChat.Text)
txtChat.SelLength = 0
txtChat.SelRTF = ThePic
Text "" & vbCrLf

Case "thepic"
If PicStarted = True And Trim(Result(2)) > "" Then ThePic = ThePic & Result(2)

Case "con" ' Connect a user
Log "Connect: " & Result(2) & vbCrLf, 32768
NewUser = Result(2)


DoEvents



Status CaptionPrefix & Result(2) & " has entered the conversation" & EndWindow

' Not yours? THen text it

If Result(4) <> sckUDP.LocalIP Then Text Result(2) & " has entered the conversation..." & vbCrLf, con, True, False, False, 8, vbCenter
DoEvents



DoEvents

' The welcome message that is displayed on each user connect

DoEvents

' Sends your username for them to add to their list.
' Rather than use Broadcast, just send it to ONE person
If NewUser <> UserName Then
'sckUDP.RemoteHost = sckUDP.RemoteHostIP

sckUDP.SendData "NChat©°®øusrø" & UserName & "ø" & MyIcon & "ø" & sckUDP.LocalIP & "ø" & NOFS
End If

' Add name to list
If List1.ListItems.Item(1).Text <> Result(2) And Result(4) <> sckUDP.LocalIP Then
List1.ListItems.Add 2, Result(4), Result(2), , ImageList1.ListImages.Item(Val(Result(3))).Key
    List1.ListItems.Item(List1.ListItems.Count).SubItems(1) = 0
End If

If frmMain.mnuDev.Visible = True And WelcomeMsg > "" Then
' If we don't replace WelcomeMSg with OWM, then
' things like +newuser+ can become stuck
OWM = WelcomeMsg
Broadcast WelcomeMsg
WelcomeMsg = OWM
End If

Case "file"

If Result(3) = UserName Then
frmBrowse.List1.AddItem Result(2)
End If

Case "lst"
If Result(2) > "" Then Log Result(2) & "'s Files are being browsed by: " & Result(3) & vbCrLf, con
If Result(2) = UserName And ShareMyFiles = True Then

For i = 0 To NOFS

If frmFileList.File1.List(i) = "" Then
Exit Sub
Else
DoEvents

Broadcast "fileø" & frmFileList.File1.List(i) & "ø" & Result(3)
DoEvents
DoEvents
DoEvents
DoEvents

End If
Next i


End If

Case "dwl"

If Result(2) = UserName Then

File2Send = frmFileList.File1.Path & "\" & Result(3)

Text Result(4) & " is downloading: " & Result(3) & vbCrLf, svr, True
Send.Show
Send.BinarySender1.Reset
Send.BinarySender1.RemoteHost = Result(5)
Send.BinarySender1.Connect
Send.BinarySender1.SendInfo

End If


Case "isr"
' Are you a real user?
Log "Is Real Check: " & Result(2) & vbCrLf, 32768
If Result(2) = UserName Then sckUDP.SendData "NChat©°®øsvrø" & UserName & "'s account is active!"

' The code to add users to your list
Case "usr"
UBR = UBound(Result)
If Result(UBR) = "True" Then RoomHost = Result(2)
If Result(2) = "" Or Result(2) = UserName Then Exit Sub

' Checks for already existing names.
    For F = 1 To List1.ListItems.Count

       If Result(2) = List1.ListItems.Item(F).Text Then
    List1.ListItems.Item(F).SubItems(1) = Result(5)

            Exit Sub
        End If

    Next F

    List1.ListItems.Add , Result(4), Result(2), , ImageList1.ListImages.Item(Val(Result(3))).Index

Case "pban"
Log Result(2) & " has been permanently banned from nchat by " & Result(3) & vbCrLf, vbRed
If Result(2) = UserName Then
MsgBox "YOU HAVE BEEN PERMANENTLY BANNED FROM NCHAT. " & UCase(Result(3)) & " HAS DECIDED THAT YOU ARE UNFIT TO PARTICIPATE IN FURTHER NCHAT DISCUSSIONS. PLEASE CONTACT GRAYDA TO DISCUSS YOUR RETURN TO NCHAT", vbCritical, "PERMANENT BAN"
Ban = True
Form_Unload (0)
End If

' When you create a new room, you become Room Admin
' This allows you to add new room admins
Case "ad1"
Log "Add Room Admin: " & Result(2) & " added by " & Result(3) & vbCrLf, vbRed
If Result(2) = UserName Then

frmAdmin.Show
Text Result(3) & " has made you a room Administrator" & vbCrLf, svr, True
Text "While a room administrator doesn't have as much power as a full" & vbCrLf, svr, True
Text "Administrator, they hold a great deal of power. Please use this power carefully" & vbCrLf, svr, True
Text "Or face being kicked or even banned from NChat!" & vbCrLf, svr, True
OldUsername = UserName
UserName = Replace(UserName, " [RA]", "")
UserName = UserName & " [RA]"
Broadcast "chuø" & OldUsername & "ø" & UserName

Status CaptionPrefix & " - You are now a Room Administrator" & EndWindow
End If

' Chucks a room admin out to the curb :P
Case "ad2"
Log "Rem Room Admin: " & Result(2) & " removed by " & Result(3) & vbCrLf, vbRed
If Result(2) = UserName Then


Unload frmAdmin
OldUsername = UserName
UserName = Replace(UserName, " [RA]", "")
Broadcast "chuø" & OldUsername & "ø" & UserName

Text "Your room Administrator rights have been removed by " & Result(3) & vbCrLf, svr, True
Status CaptionPrefix & " - Your admin rights have been removed" & EndWindow
End If

Case "pm1"
' Concerned about privacy? Comment out this line.
' It stops people from listening in to Private
' conversations

' It's just a way for me to test Private Messages.
Log "PM: " & Result(4) & " (to " & Result(2) & "): " & Result(3) & vbCrLf, vbBlack, True

' Is it for you?
If Result(3) = "" Then Exit Sub
LastFrom = UserName
If Result(2) = UserName Then 'And Result(4) <> Username Then
Randomize

' If this is false, then don't show the box.
' This only changes after you recieve the
' first message.
If NewMessage = False And Me.Visible = False Then
Tray.Box Result(4) & " has sent you a private message! Double click this icon to read it!", "New Private Message"
NewMessage = True
End If

For i = 1 To lstIgnore.ListItems.Count
'If lstIgnore.ListItems.Item(i) = "" Then Exit For
If lstIgnore.ListItems.Item(i) = Result(4) Then
Text Result(4) & " (Who is on your ignore list), tried to send you a message", svr, True
DoAutoBot sData
Exit Sub
End If
Next i

For i = LBound(CW) To UBound(CW)
If CW(i).Tag = Result(4) And CW(i).Visible = True Then
Txt2 Result(4) & " ::  " & Result(3) & vbCrLf, act, Int(i)
CheckRTF CW(i).Text1
DoAutoBot sData
Exit Sub
End If
Next i


If ListView1.Visible = False Then Text "You have a new Private Message from " & Result(4) & vbCrLf, svr, True
ListView1.ListItems.Add , "R…" & Result(3) & "…" & Int(Rnd * 5000), Result(4), , ImageList1.ListImages.Item(ImageList1.ListImages.Count - 2).Key
' Got a Away Message set? Display it then

If AwayMSG <> "" And AwayMessage = True And Result(5) = "AutoAway" And Result(5) = "" Then
Broadcast "pm1+d+" & Result(4) & "+d+" & AwayMSG & "+d+" & UserName & "+d+AutoAway"
End If

End If

' Disconnects a user from NChat and removes their name from
' the list of users
Case "dis"
Log "Disconnect: " & Result(2) & vbCrLf, vbRed
Text Result(2) & " has left the room!" & vbCrLf, dis, True, , , , vbCenter
Status CaptionPrefix & Result(2) & " has left the room!!" & EndWindow

' The DoAutoBot(sData) sub is called here, because there
' is an exit sub in here as well, and DoAutoBot(sData)
' can't be called. This also happens with "con"
DoAutoBot sData
' Scroll through the list and removes the user in question
For i = 1 To List1.ListItems.Count
DoEvents
If List1.ListItems.Item(i).Text = Result(2) Then
List1.ListItems.Remove List1.ListItems.Item(i).Index

Exit Sub
End If
DoEvents
Next i


' Gets rid of an admin's rights
' Unless you are a "True" admin
Case "rem"
Log "Kill Admin: " & Result(2) & " killed by " & Result(3) & vbCrLf, vbRed
If Result(2) = UserName And TrueAdmin = False Then
mnuDev.Visible = False
Unload frmAdminWindow
mnuUE.Visible = False
OldUsername = UserName
UserName = Replace(UserName, " [A]", "")
Broadcast "chuø" & OldUsername & "ø" & UserName
Text Result(3) & " has taken away your admin rights!" & vbCrLf, svr, True

ElseIf Result(2) = UserName And TrueAdmin = True Then
Text Result(3) & " tried to take your admin rights! What a rat!!" & vbCrLf, svr, True
End If

Case "move"
' This case deals with the whiteboard.
' The line colour, location, and size are sent in one
' packet, to save some network bandwidth (pfft. Yeah right. To draw
' a 3cm line, takes about 50 move commands...). The board
' should really have it's own winsock control, and be private
' between the 2 people, but I can't be stuffed...

'Log "Move Cursor: " & Result(2) & " from " & Result(3) & " X: " & Result(4) & " Y: " & Result(5) & vbCrLf
If Result(2) = UserName Then
For i = LBound(CW) To UBound(CW)

If CW(i).Tag = Result(3) Then
    With CW(i)
        .Picture1.DrawWidth = Result(7)
        .Picture1.Enabled = False
        .Picture1.Line (lastX, lastY)-(Result(4), Result(5)), Result(6)

        lastX = Result(4)
        lastY = Result(5)
        .Picture1.Enabled = True
    End With
End If

Next i
End If


Case "fill"
Log "Fill " & Result(2) & "'s Picture Box (From " & Result(3) & ") at points X:" & Result(4) & " Y:" & Result(5) & " with the colour " & Result(6), vbRed, True

If Result(2) = UserName Then
For i = LBound(CW) To UBound(CW)

If CW(i).Tag = Result(3) Then
'ExtFloodFill CW(i).Picture1.hDC, Result(4), Result(5), Result(6), 1
CW(i).Picture1.FillColor = Result(6)
ExtFloodFill CW(i).Picture1.hdc, Result(4), Result(5), CW(i).Picture1.point(Result(4), Result(5)), 1
End If
Next i
End If



Case "Clear"
'Log Result(2) & "'s Whiteboard has been cleared by " & Result(3)
If Result(2) = UserName Then
For i = LBound(CW) To UBound(CW)

If CW(i).Tag = Result(3) Then
            CW(i).Picture1.Cls
                       Exit Sub
End If


Next i
End If





' Change your username and your Icon on all user lists
Case "chu"
Log "Change Username: " & Result(2) & " to " & Result(3) & vbCrLf, 32768

For i = 1 To List1.ListItems.Count

If List1.ListItems.Item(i).Text = Result(2) Then
List1.ListItems.Item(i).Text = Result(3)
List1.ListItems.Item(i).Icon = ImageList1.ListImages.Item(Result(4)).Key
DoAutoBot sData
Exit Sub
End If
Next i
DoEvents
DoEvents

' Makes a really big heading to notify of events, crush people
' or just for fun
Case "hea"
Log "Heading: " & Result(2) & " (" & Result(3) & ")" & vbCrLf, vbBlue
Text Result(2) & vbCrLf, Heading, True, False, False, 18, vbCenter
Text "" & vbCrLf, Msg, False, False, False, txtChat.Font.Size


' when you purchase a kick user item from the store,
' it kicks someone and says "<USERNAME> has kicked you from the
' room", rather than saying that the admin did it
Case "kun"
Log Result(3) & " kicked " & Result(2) & vbCrLf, vbRed
If Result(2) = UserName Or Result(2) = sckUDP.LocalIP And TrueAdmin = False Then

MsgBox Result(3) & " has kicked you from the NChat Chatrooms", vbCritical, "Kicked by User"
mnuDev.Visible = False
If TrueAdmin = False Then
Form_Unload (0)
Broadcast "disø" & UserName
Else
Text Result(3) & " tried to kick you from NChat, what a rat!!" & vbCrLf, svr, True
End If
End If

' Another kind of kick, this time from the admin
Case "ksv"
Log Result(3) & " (Admin) kicked " & Result(2) & vbCrLf, vbRed
If Result(2) = UserName Or Result(2) = sckUDP.LocalIP And TrueAdmin = False Then

'mnuDev.Visible = False

If TrueAdmin = False Then
Broadcast "disø" & UserName

Broadcast "svrø+username+ has been kicked from NChat"
MsgBox "The administrator (" & Result(3) & ") has kicked you from the room. Please correct your behaviour before re-connecting.", vbCritical, "Kicked by Administrator"


Form_Unload (0)
Else
Text Result(3) & " tried to kick you from NChat, what a rat!!", svr, True
End If

End If

Case "pic"
' Experimental... LARGE amounts of data sent, that's why it's experimental
Result(2) = Replace(Result(2), "##", "{\")
txtChat.SelStart = Len(txtChat.Text) - 1
txtChat.SelLength = 1
 txtChat.SelRTF = txtChat.SelRTF & Result(2)
 
Case "snd"
Log Result(4) & " sent " & Result(2) & " " & Result(3) & " NCredits" & vbCrLf, ThatsGood
' Sends NCredits to a user
' Result(2) = Username
' Result(3) = How Many NCredits
' Result(4) = Who are the NCredits from?

If Result(2) = UserName Then

NCredits = NCredits + Result(3)
Text Result(4) & " has given you " & Result(3) & " NCredits!" & vbCrLf, ThatsGood, True
Status CaptionPrefix & " - Someone has given you " & Result(3) & " NCredits!" & EndWindow
End If

Case "force"
If Result(2) = UserName Then
MsgBox "Because of your actions, you have been kicked from NChat. Please correct your behaviour before re-entering!", vbCritical, "Force Kicked"
mnuDev.Visible = False
TrueAdmin = False
Form_Unload (0)
End If

' Prints a user's statistics
Case "pip"
Log Result(2) & "'s Stats Requested by " & Result(3) & vbCrLf, 32768
If Result(2) = UserName Then
Broadcast "svrø" & UserName & "'s statistics (" & Time & "):+newline++newline+IP: " & sckUDP.LocalIP & "+newline+NCredits: " & NCredits & "+newline+Last Message from: " & sckUDP.RemoteHostIP & "+newline+Time on NChat: " & NChatTime & " seconds+newline+Core Version: " & App.Major & "/" & App.Minor & "/" & App.Revision & "+newline+Smileys on?: " & Smiley & "+newline+Real Username: " & GetUserName & "+newline+Admin: " & mnuDev.Visible & "+newline+Swearing on?: " & Swearing & "+newline+Total Icons: " & TotalIcons & "+newline+Profile: " & IniFile(1) & "+newline+On Computer: " & Environ("Computername") & "+newline+Files Shared?: " & ShareMyFiles & "+newline+Number of files shared: " & NOFS
Status CaptionPrefix & " - Your statistics requested" & EndWindow
End If

' Change your userlist icon
Case "chi"
Log Result(2) & " changed icon to index " & Result(3) & vbCrLf, 32768
For i = 1 To List1.ListItems.Count
If List1.ListItems.Item(i).Text = Result(2) Then List1.ListItems.Item(i).SmallIcon = ImageList1.ListImages.Item(Val(Result(3))).Key
Next i

' Redirect a user to another room
Case "red"
Log Result(3) & " redirected to: " & Result(4) & " (" & Result(2) & ") by " & Result(5) & vbCrLf, vbRed
If IsInternet = True Then
' Click the disconnect menu
mnuDisconnect_Click
MsgBox "You have been redircted from the NChat server. You are now connected to NChat via the NETWORK. To re-connect to the NChat server, open the Main Menu", vbCritical, "You have been redirected"
Exit Sub
End If

If Result(3) = UserName Then
Text "You have been redirected to room #" & Result(2) & " (" & Result(4) & ")" & vbCrLf, ThatsBad, True
NewRoom Result(2), Result(4)
End If

' Lets me know if people are ghosting others >:D
Case "fak"
Log "GHOST: " & Result(2) & " - " & Result(3) & " (" & Result(4) & ")" & vbCrLf, vbRed
Text StartMSG & " " & Result(2) & " " & EndMSG & " " & Result(3) & vbCrLf, Msg

' Changes the caption to say the new message
Status CaptionPrefix & Result(2) & " - " & Result(3) & EndWindow

Case Else
' Not any of the above? Treat it as unknown orF
' Raw Data message
If Result(1) = UserName Or Result(0) = "room" Then Exit Sub

Text Result(1) & vbCrLf, Msg
Log "Unknown Command / Raw Data: " & Result(1) & vbCrLf, vbBlue
End Select

SB1.Panels(3).Picture = picGreen.Picture
If txtChat.SelLength = 0 Then txtChat.SelStart = Len(txtChat.Text)


' Scans through the latest batch of text for
' a smiley :) :evil
CheckRTF frmMain.txtChat
StartAt = 1

DoAutoBot sData

txtChat.SelStart = Len(txtChat.Text)

' Resets the RemoteHost
' So all messages can reach their destinations
sckUDP.RemoteHost = Address

End Sub


Private Sub sckUDP_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
If IsInternet = True Then MsgBox "There was an error with your NChat. Please report the following details to Grayda (www.solidinc.tk)" & vbCrLf & vbCrLf & Number & vbCrLf & Description, vbCritical, "NChat Error"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Timer2_Timer
' DateTime  : 3/12/2004 18:23
' Author    : Grayda
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Timer2_Timer()
   On Error GoTo Timer2_Timer_Error

On Error Resume Next
'MessageColour = Msg
If ListView1.ListItems.Count = 0 Then

Tray.IconHandle = Me.Icon
Else
Tray.IconHandle = ImageList1.ListImages.Item(ImageList1.ListImages.Count - 2).Picture
End If

' The room you are in and how long you have been on NChat for
NChatTime = NChatTime + 1
SB1.Panels(2).Text = "You are currently in room: " & RoomName & " (#" & RoomID & ")"

If CreatedRoom = True Then
RoomTime = RoomTime + 1
RoomBroadcast "roomø" & "R" & sckUDP.LocalPort & "ø" & RoomName & "ø" & MyIcon
hr5.Visible = True
mnuChangeRoomInfo.Visible = True
Else
hrh.Visible = False
mnuChangeRoomInfo.Visible = False
End If

If ShareMyFiles = False Then
NOFS = 0
Else
NOFS = frmFileList.File1.ListCount
End If

' Sends your online presence to everyone. THere
' are better ways to do this, but I can't be
' bothered :P
Broadcast "usrø" & UserName & "ø" & MyIcon & "ø" & sckUDP.LocalIP & "ø" & NOFS & "ø" & CreatedRoom
List1.ListItems.Item(1).SubItems(1) = NOFS
If FilePattern > "" Then frmFileList.File1.Pattern = FilePattern

   On Error GoTo 0
   Exit Sub

Timer2_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Timer2_Timer of Form frmMain"

End Sub

Private Sub txtChat_Change()
SetAutoURL4RTB txtChat
End Sub

Private Sub txtChat_Click()
' This code ISN'T mine, so here is the copyright info:

' This OCX + the Source code is in fact "Copyrighted" however with an exception
' You may use this ocx + source code in your program as long as you
' mention leave the name of these controls the same in your program.
'(ie URTB_V) 'however when you use it in design time ...you can change
' The reason for this is... so if programmers use API SPY... they can see it's mine
'UDMX IOCP© : Bringing you the source"
'''MY INFO'''
'Net NAME: UDMX IOCP©
'AGE: 15
'VB PROGRAMMING LEVEL: ADV/PRO
'Current Project: PSC Chat Deluxe© (THIS IS FOR INTRADREAM.COM)
'LOCATION: US/CA
'Other notes: having something you need help with??? If yes Send me a message on
'AIM: UDMX IOCP
'MSN MESSENGER: UDMX_IOCP@HOTMAIL.COM
'ICQ: grrr I hate this software so don't have one
'PSC Chat: UDMXIOCP ... UDMX  'to download PSC Chat created by
'   Tim (a IntraDream© member like me) and some other ppl go to www.IntraDream.com
'
    Dim lngRet 'use ARRAY IT'S FASTER.. one for left... the other for right ...not hard
    Dim i As Long
    Dim ii As Long
    LeftDat() = SplitVB5(sFrom3Left, "*")
          Dim strString As String

       strString = sFrom3Right2 & "*" & sFrom3Right3 & "*" & sFrom3Right4 & "*" & sFrom3Right5 & "*" & sFrom3Right6

    RightDat() = SplitVB5(strString, "*")
    For i& = 0 To UBound(LeftDat()) 'we'll use lcase
        If Left(LCase(htxt), Len(LeftDat(i&)) + 3) = (LeftDat(i&) & "://") Or Left(LCase(htxt), 4) = "www." Then 'www is special hehehe
            For ii& = 0 To UBound(RightDat())
                If Right(LCase(htxt), Len(RightDat(ii&)) + 1) = ("." & RightDat(ii&)) Then
                    lngRet = ShellExecute(0&, "Open", htxt, "", vbNullString, SW_SHOWNORMAL)
                    Exit Sub
                End If
            Next ii&
        End If
    Next i&
    If InStr(htxt, "@") = 1 Then
        htxt = Replace(htxt, "mailto:", "")
  lngRet = ShellExecute(0&, "Open", "mailto:" + htxt, "", vbNullString, SW_SHOWNORMAL)
    End If
End Sub

Private Sub txtChat_KeyPress(KeyAscii As Integer)
' Not trying to copy anything?
' then set the focus on the send box, with
' your typed key
If txtChat.SelLength = 0 Then
txtSend.SetFocus
txtSend.Text = txtSend.Text & Chr(KeyAscii)
txtSend.SelStart = Len(txtSend.Text)
End If
End Sub

Private Sub txtChat_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error Resume Next

    htxt = GetHyperlink(txtChat, X, y)
End Sub

Private Sub txtSend_KeyDown(KeyCode As Integer, Shift As Integer)
' Just some general shortcuts
Select Case KeyCode
Case vbKeyReturn
cmdSend_Click
Case vbKeyF2
frmSmileys.Show
Case vbKeyF3
frmStore.Show
Case vbKeyF5
frmNewProfile.Show
Case vbKeyF6
frmOptions.Show
Case vbKeyF12
mnuHide_Click
End Select
End Sub



Private Sub ul1_Click()
' Send a private message from the user list
On Error Resume Next


For i = LBound(CW) To UBound(CW)
If CW(i).Tag = "" Then
CW(i).WindowState = 0
CW(i).Show
CW(i).Tag = List1.SelectedItem.Text
CW(i).Picture1.Tag = List1.SelectedItem.Key
Exit Sub
End If
Next i
End Sub

Private Sub ul2_Click()
On Error Resume Next
' Sends some NCredits to the person you clicked on
' on the user list
ToWho = List1.SelectedItem.Text
HowMany = InputBox("How many NCredits do you want to give?", "Give NCredits")
If NCredits > HowMany And HowMany > 0 Then
Broadcast ("sndø" & ToWho & "ø" & HowMany & "ø" & UserName)
NCredits = NCredits - HowMany
ElseIf HowMany <> "" Then
Text "You do not have enough NCredits to give away!!" & vbCrLf, ThatsBad, True
ElseIf HowMany = "" Then
Exit Sub
End If

End Sub

Private Sub ul4_Click()
MsgBox "WARNING: This file transfer system is highly unstable. Some files may not be listed, or transfered. To see the entire list, click the button on the file list a few times, until you get some files", vbCritical, "Transfer not yet complete!"
frmBrowse.Tag = List1.SelectedItem.Text
frmBrowse.List1.Tag = List1.SelectedItem.Key
frmBrowse.Show

DoEvents
DoEvents
Broadcast "lstø" & frmBrowse.Tag & "ø" & UserName

End Sub
