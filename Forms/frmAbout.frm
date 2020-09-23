VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About NChat Alpha Generation 2"
   ClientHeight    =   4875
   ClientLeft      =   -1395
   ClientTop       =   -780
   ClientWidth     =   6915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   30500
      Left            =   0
      ScaleHeight     =   30495
      ScaleWidth      =   6975
      TabIndex        =   0
      Top             =   0
      Width           =   6975
      Begin VB.Timer Timer2 
         Interval        =   1
         Left            =   720
         Top             =   2880
      End
      Begin VB.Timer Timer1 
         Interval        =   50
         Left            =   240
         Top             =   2880
      End
      Begin VB.Line Line11 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   1560
         X2              =   5280
         Y1              =   30120
         Y2              =   30120
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "firestorm_visual@hotmail.com"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   1612
         MouseIcon       =   "frmAbout.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   51
         Top             =   29760
         Width           =   3690
      End
      Begin VB.Label Label50 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAbout.frx":0352
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1425
         Left            =   120
         TabIndex        =   50
         Top             =   28200
         Width           =   6540
      End
      Begin VB.Label Label49 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAbout.frx":041D
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2265
         Left            =   195
         TabIndex        =   49
         Top             =   25920
         Width           =   6540
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   1410
         X2              =   5490
         Y1              =   25440
         Y2              =   25440
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.planetsourcecode.com"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   1372
         MouseIcon       =   "frmAbout.frx":053E
         MousePointer    =   99  'Custom
         TabIndex        =   48
         Top             =   25080
         Width           =   4170
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PLANET SOURCE CODE"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2272
         TabIndex        =   47
         Top             =   24720
         Width           =   2370
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   1830
         X2              =   5070
         Y1              =   24480
         Y2              =   24480
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "http://balz.stormynight.net"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   1845
         MouseIcon       =   "frmAbout.frx":0890
         MousePointer    =   99  'Custom
         TabIndex        =   46
         Top             =   24120
         Width           =   3225
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HEYALLO'S HAVEN - FORUMS"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1965
         TabIndex        =   45
         Top             =   23760
         Width           =   2985
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   1560
         X2              =   5400
         Y1              =   23520
         Y2              =   23520
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "http://firestorm.stormynight.net"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   1522
         MouseIcon       =   "frmAbout.frx":0BE2
         MousePointer    =   99  'Custom
         TabIndex        =   44
         Top             =   23160
         Width           =   3870
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FIRESTORM VISUAL SOFTWARE - MY SITE"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1290
         TabIndex        =   43
         Top             =   22800
         Width           =   4335
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DON'T FORGET TO VISIT"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1890
         TabIndex        =   42
         Top             =   21240
         Width           =   3135
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   2130
         X2              =   4770
         Y1              =   22560
         Y2              =   22560
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.solidinc.tk"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   2130
         MouseIcon       =   "frmAbout.frx":0F34
         MousePointer    =   99  'Custom
         TabIndex        =   41
         Top             =   22200
         Width           =   2655
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SOLID INC. MEDIA PRODUCTIONS - MY SITE"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1170
         TabIndex        =   40
         Top             =   21840
         Width           =   4575
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PLUS THE R.I.S.K TEAM"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2250
         TabIndex        =   39
         Top             =   20640
         Width           =   2415
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AND THE SOLID INC. REPS"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2100
         TabIndex        =   38
         Top             =   20280
         Width           =   2715
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HUTCHY"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3060
         TabIndex        =   37
         Top             =   19920
         Width           =   915
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NIEUMI"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3045
         TabIndex        =   36
         Top             =   19560
         Width           =   780
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DR. IDGE"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3015
         TabIndex        =   35
         Top             =   19200
         Width           =   930
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GREETINGS TO:"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2445
         TabIndex        =   34
         Top             =   18720
         Width           =   2025
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AND ALL THE OTHER PEOPLE"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2002
         TabIndex        =   33
         Top             =   18120
         Width           =   2910
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TRAVIS LETHBORG"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2475
         TabIndex        =   32
         Top             =   17760
         Width           =   1965
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GRAYDA"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3022
         TabIndex        =   31
         Top             =   17400
         Width           =   870
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HEYALLO"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3000
         TabIndex        =   30
         Top             =   17040
         Width           =   915
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DA JEDAZ"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2970
         TabIndex        =   29
         Top             =   16680
         Width           =   975
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BETA TESTERS"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2490
         TabIndex        =   28
         Top             =   16200
         Width           =   1935
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   2190
         X2              =   4710
         Y1              =   15960
         Y2              =   15960
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.allapi.net"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   2182
         MouseIcon       =   "frmAbout.frx":1286
         MousePointer    =   99  'Custom
         TabIndex        =   27
         Top             =   15600
         Width           =   2550
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VARIOUS FUNCTIONS - KPD-TEAM"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1672
         TabIndex        =   26
         Top             =   15240
         Width           =   3570
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=55313&lngWId=1"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   135
         MouseIcon       =   "frmAbout.frx":15D8
         MousePointer    =   99  'Custom
         TabIndex        =   25
         Top             =   14760
         Width           =   6690
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   120
         X2              =   6855
         Y1              =   15015
         Y2              =   15000
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AUTHOR UNKNOWN - REALLY COOL SPLASH SCREEN"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   750
         TabIndex        =   24
         Top             =   14400
         Width           =   5400
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=11235&lngWId=1"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   120
         MouseIcon       =   "frmAbout.frx":192A
         MousePointer    =   99  'Custom
         TabIndex        =   23
         Top             =   13920
         Width           =   6675
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   105
         X2              =   6840
         Y1              =   14175
         Y2              =   14160
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AUTHOR UNKNOWN - VB6 TO VB5 FUNCTIONS"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1050
         TabIndex        =   22
         Top             =   13560
         Width           =   4800
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   1770
         X2              =   5130
         Y1              =   13320
         Y2              =   13320
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.IntraDream.com"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   1777
         MouseIcon       =   "frmAbout.frx":1C7C
         MousePointer    =   99  'Custom
         TabIndex        =   21
         Top             =   12960
         Width           =   3360
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UDMX IOCP - HYPERLINKS AND SMILEY CODE"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1080
         TabIndex        =   20
         Top             =   12600
         Width           =   4755
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "http://pcseries.sourceforge.net/pscode/BinaryTransferControl.zip"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   345
         MouseIcon       =   "frmAbout.frx":1FCE
         MousePointer    =   99  'Custom
         TabIndex        =   19
         Top             =   12120
         Width           =   6210
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   360
         X2              =   6600
         Y1              =   12420
         Y2              =   12420
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NCHAT USES CODES FROM"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1695
         TabIndex        =   18
         Top             =   11280
         Width           =   3525
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "KENNY LAI - BINARY TRANSFER"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1845
         TabIndex        =   17
         Top             =   11760
         Width           =   3210
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   2160
         X2              =   4680
         Y1              =   10800
         Y2              =   10800
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.trillian.cc"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   2175
         MouseIcon       =   "frmAbout.frx":2320
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Top             =   10440
         Width           =   2550
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(:brb :cd and :monkey FROM TRILLIAN)"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1335
         TabIndex        =   15
         Top             =   10080
         Width           =   4230
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(EXCEPT :solid :nchat and :| BY GRAYDA)"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1245
         TabIndex        =   14
         Top             =   9600
         Width           =   4410
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "JAYANTH KUMAR J"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2490
         TabIndex        =   13
         Top             =   9240
         Width           =   1920
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SMILEYS CREATED BY"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2010
         TabIndex        =   12
         Top             =   8760
         Width           =   2880
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GRAYDA"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3015
         TabIndex        =   11
         Top             =   8160
         Width           =   870
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GRAPHICS  DRAWN BY:"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1905
         TabIndex        =   10
         Top             =   7680
         Width           =   3105
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GRAYDA"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3015
         TabIndex        =   9
         Top             =   6960
         Width           =   870
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "DESIGNED  BY:"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2430
         TabIndex        =   8
         Top             =   6480
         Width           =   2055
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FIRESTORM VISUAL SOFTWARE"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1815
         TabIndex        =   7
         Top             =   5760
         Width           =   3285
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SOLID INC. AND"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   2880
         TabIndex        =   6
         Top             =   5040
         Width           =   1155
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "GRAYDA OF"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   2985
         TabIndex        =   5
         Top             =   4320
         Width           =   945
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "CODED BY:"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2730
         TabIndex        =   4
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "GENERATION II"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2430
         TabIndex        =   3
         Top             =   3120
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "ALPHA"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2730
         TabIndex        =   2
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "NCHAT"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   48
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   2190
         TabIndex        =   1
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Image Image1 
         Height          =   975
         Left            =   2910
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' OK this form is our about form. All controls
' are contained in a picture box, so we can scroll
' the whole thing without too much code.

' I know there is a scrollhdc API, but this is simpler
' and scrolls at a rate that stops most flickering

' Get the cursor position, so we can detect
' if it's within our form
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

' This works out the position of our control
' on the screen. This works hand-in-hand with
' the functions above.
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
' Dim our R1 as a new rect, or you
' can get a byref error. I sat here for an hour
' trying to work out what the byref error was.
' Turns out I forgot to dim my rect \:)
Dim R1 As RECT
Dim P1 As POINTAPI

Private Sub Form_Load()
' Load our picture for the start of the about box
Image1.Picture = frmSmileys.imgIcon(107).Picture
End Sub

Private Sub Label16_Click()
' RunHyper is a hyperlink runner.
' Saves on code :)
RunHyper Label6.Caption

End Sub


Private Sub Label19_Click()
RunHyper Label9.Caption

End Sub

Private Sub Label21_Click()
RunHyper Label21.Caption
End Sub

Private Sub RunHyper(Hyperlink As String)
' Run our hyperlink using an extended shell command
lngRet = ShellExecute(0&, "Open", Hyperlink, "", vbNullString, SW_SHOWNORMAL)
End Sub

Private Sub Label23_Click()
RunHyper Label23

End Sub

Private Sub Label27_Click()
RunHyper Label27
End Sub

Private Sub Label41_Click()
RunHyper Label41
End Sub

Private Sub Label44_Click()
RunHyper Label44
End Sub

Private Sub Label46_Click()
RunHyper Label46
End Sub

Private Sub Label51_Click()
' oooh! A Special link. Sends an e-mail to me
RunHyper "mailto:" & Label51
End Sub

Private Sub Timer1_Timer()
' Slowly scrolls our credits
' Doesn't flicker on my computer, not sure about others
Picture1.Top = Picture1.Top - 50

If Picture1.Top < -Picture1.Height - Me.Height Then Picture1.Top = Me.Height


End Sub

Private Sub Timer2_Timer()
Dim R1 As RECT
Dim P1 As POINTAPI
' Get the location of our form
GetWindowRect Me.hwnd, R1
' Get the location of our cursor
GetCursorPos P1

' Is our cursor over our form? If so, then stop scrolling
If P1.X < R1.Right And P1.X > R1.Left And P1.y < R1.Bottom And P1.y > R1.Top Then
Timer1.Enabled = False
Else
Timer1.Enabled = True
End If

End Sub
