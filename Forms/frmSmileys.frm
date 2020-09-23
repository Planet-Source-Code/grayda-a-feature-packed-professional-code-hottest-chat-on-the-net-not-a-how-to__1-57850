VERSION 5.00
Begin VB.Form frmSmileys 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Smileys that you can use in NChat"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7590
   Icon            =   "frmSmileys.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton NChat_Button1 
      Caption         =   "Save all Smileys"
      Height          =   375
      Left            =   2760
      TabIndex        =   62
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Index           =   109
      Left            =   240
      Picture         =   "frmSmileys.frx":2CFA
      Tag             =   "(*)"
      Top             =   3240
      Width           =   240
   End
   Begin VB.Label Label17 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "(*)"
      Height          =   255
      Left            =   255
      TabIndex        =   64
      Top             =   3555
      Width           =   255
   End
   Begin VB.Image imgIcon 
      Height          =   420
      Index           =   108
      Left            =   5640
      Picture         =   "frmSmileys.frx":3084
      Tag             =   ":afro"
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":afro"
      Height          =   255
      Left            =   5640
      TabIndex        =   63
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":nchat"
      Height          =   255
      Left            =   6360
      TabIndex        =   61
      Top             =   3480
      Width           =   615
   End
   Begin VB.Image imgIcon 
      Height          =   840
      Index           =   107
      Left            =   6360
      Picture         =   "frmSmileys.frx":316E
      Tag             =   ":nchat"
      Top             =   2640
      Width           =   840
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   106
      Left            =   4800
      Picture         =   "frmSmileys.frx":3B6C
      Stretch         =   -1  'True
      Tag             =   ":solid"
      Top             =   2640
      Width           =   600
   End
   Begin VB.Label Label18 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":solid"
      Height          =   255
      Left            =   4800
      TabIndex        =   60
      Top             =   3120
      Width           =   615
   End
   Begin VB.Image imgIcon 
      Height          =   345
      Index           =   105
      Left            =   3960
      Picture         =   "frmSmileys.frx":4113
      Tag             =   ":lick"
      Top             =   2670
      Width           =   690
   End
   Begin VB.Label Label20 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":lick"
      Height          =   255
      Left            =   3960
      TabIndex        =   59
      Top             =   3030
      Width           =   615
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   47
      Left            =   6000
      Picture         =   "frmSmileys.frx":4230
      Tag             =   ":|"
      Top             =   2160
      Width           =   225
   End
   Begin VB.Label Label19 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":|"
      Height          =   255
      Left            =   6000
      TabIndex        =   58
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "-->"
      Height          =   255
      Left            =   3480
      TabIndex        =   57
      Top             =   3045
      Width           =   255
   End
   Begin VB.Image imgIcon 
      Height          =   210
      Index           =   104
      Left            =   3480
      Picture         =   "frmSmileys.frx":4298
      Tag             =   "-->"
      Top             =   2760
      Width           =   285
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "<--"
      Height          =   255
      Left            =   3000
      TabIndex        =   56
      Top             =   3045
      Width           =   255
   End
   Begin VB.Image imgIcon 
      Height          =   210
      Index           =   103
      Left            =   3000
      Picture         =   "frmSmileys.frx":4301
      Tag             =   "<--"
      Top             =   2760
      Width           =   285
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":monkey"
      Height          =   255
      Left            =   2280
      TabIndex        =   55
      Top             =   3045
      Width           =   615
   End
   Begin VB.Image imgIcon 
      Height          =   255
      Index           =   102
      Left            =   2400
      Picture         =   "frmSmileys.frx":436A
      Tag             =   ":monkey"
      Top             =   2760
      Width           =   270
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":cd"
      Height          =   255
      Left            =   1920
      TabIndex        =   54
      Top             =   3045
      Width           =   375
   End
   Begin VB.Image imgIcon 
      Height          =   300
      Index           =   101
      Left            =   1920
      Picture         =   "frmSmileys.frx":45D5
      Tag             =   ":cd"
      Top             =   2760
      Width           =   300
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":brb"
      Height          =   255
      Left            =   1440
      TabIndex        =   53
      Top             =   3045
      Width           =   375
   End
   Begin VB.Image imgIcon 
      Height          =   270
      Index           =   100
      Left            =   1440
      Picture         =   "frmSmileys.frx":4A0D
      Tag             =   ":brb"
      Top             =   2760
      Width           =   255
   End
   Begin VB.Image imgIcon 
      Height          =   285
      Index           =   99
      Left            =   840
      Picture         =   "frmSmileys.frx":4B74
      Tag             =   ":weed"
      Top             =   2760
      Width           =   285
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":weed"
      Height          =   255
      Left            =   720
      TabIndex        =   52
      Top             =   3045
      Width           =   495
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Index           =   98
      Left            =   240
      Picture         =   "frmSmileys.frx":4F72
      Tag             =   ":finger"
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":cry"
      Height          =   255
      Left            =   240
      TabIndex        =   51
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":idiot"
      Height          =   255
      Left            =   7080
      TabIndex        =   50
      Top             =   2400
      Width           =   495
   End
   Begin VB.Image imgIcon 
      Height          =   285
      Index           =   97
      Left            =   7080
      Picture         =   "frmSmileys.frx":5035
      Tag             =   ":idiot"
      Top             =   2115
      Width           =   285
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":finger"
      Height          =   255
      Left            =   6480
      TabIndex        =   49
      Top             =   2400
      Width           =   495
   End
   Begin VB.Image imgIcon 
      Height          =   300
      Index           =   96
      Left            =   6480
      Picture         =   "frmSmileys.frx":54BC
      Tag             =   ":finger"
      Top             =   2115
      Width           =   300
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSmileys.frx":5634
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   48
      Top             =   3840
      Width           =   7335
   End
   Begin VB.Label Label97 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":toiletclaw"
      Height          =   255
      Left            =   120
      TabIndex        =   47
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label96 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":guns"
      Height          =   255
      Left            =   3840
      TabIndex        =   46
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label95 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":guitar"
      Height          =   255
      Left            =   6120
      TabIndex        =   45
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label94 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":light"
      Height          =   255
      Left            =   4320
      TabIndex        =   44
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label93 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":ass"
      Height          =   255
      Left            =   600
      TabIndex        =   43
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label92 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":wave"
      Height          =   255
      Left            =   3720
      TabIndex        =   42
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label91 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":bala"
      Height          =   255
      Left            =   600
      TabIndex        =   41
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label90 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":king"
      Height          =   255
      Left            =   120
      TabIndex        =   40
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label89 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":heart"
      Height          =   255
      Left            =   2160
      TabIndex        =   39
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label88 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":tilt"
      Height          =   255
      Left            =   3240
      TabIndex        =   38
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label87 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":alien"
      Height          =   255
      Left            =   6480
      TabIndex        =   37
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label86 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":devil"
      Height          =   255
      Left            =   7080
      TabIndex        =   36
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label85 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":angel"
      Height          =   195
      Left            =   4560
      TabIndex        =   35
      Top             =   1800
      Width           =   435
   End
   Begin VB.Label Label84 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":erm"
      Height          =   255
      Left            =   2640
      TabIndex        =   34
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label83 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":flush"
      Height          =   255
      Left            =   2280
      TabIndex        =   33
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label82 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":lol"
      Height          =   255
      Left            =   240
      TabIndex        =   32
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label Label81 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":mad"
      Height          =   255
      Left            =   1800
      TabIndex        =   31
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label80 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":gg"
      Height          =   255
      Left            =   2880
      TabIndex        =   30
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label Label79 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":poo"
      Height          =   255
      Left            =   6840
      TabIndex        =   29
      Top             =   1800
      Width           =   375
   End
   Begin VB.Image imgIcon 
      Height          =   345
      Index           =   95
      Left            =   5280
      Picture         =   "frmSmileys.frx":56C0
      Tag             =   ":shoot"
      Top             =   1440
      Width           =   1515
   End
   Begin VB.Image imgIcon 
      Height          =   375
      Index           =   94
      Left            =   6480
      Picture         =   "frmSmileys.frx":5BDD
      Tag             =   ":alien"
      Top             =   0
      Width           =   270
   End
   Begin VB.Image imgIcon 
      Height          =   315
      Index           =   93
      Left            =   4560
      Picture         =   "frmSmileys.frx":5C7A
      Tag             =   ":angel"
      Top             =   1440
      Width           =   600
   End
   Begin VB.Image imgIcon 
      Height          =   330
      Index           =   92
      Left            =   120
      Picture         =   "frmSmileys.frx":5DC8
      Tag             =   ":king"
      Top             =   720
      Width           =   345
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Index           =   91
      Left            =   6840
      Picture         =   "frmSmileys.frx":617C
      Tag             =   ":poo"
      Top             =   1440
      Width           =   240
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   90
      Left            =   1800
      Picture         =   "frmSmileys.frx":64F2
      Tag             =   ":mad"
      Top             =   1440
      Width           =   255
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   89
      Left            =   240
      Picture         =   "frmSmileys.frx":65C9
      Tag             =   ":lol"
      Top             =   1440
      Width           =   255
   End
   Begin VB.Image imgIcon 
      Height          =   300
      Index           =   88
      Left            =   5400
      Picture         =   "frmSmileys.frx":6696
      Tag             =   ":blah"
      Top             =   720
      Width           =   660
   End
   Begin VB.Image imgIcon 
      Height          =   360
      Index           =   87
      Left            =   2880
      Picture         =   "frmSmileys.frx":67B0
      Tag             =   ":beat"
      Top             =   2040
      Width           =   525
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   86
      Left            =   4320
      Picture         =   "frmSmileys.frx":6BB3
      Tag             =   ":light"
      Top             =   720
      Width           =   225
   End
   Begin VB.Image imgIcon 
      Height          =   300
      Index           =   85
      Left            =   3840
      Picture         =   "frmSmileys.frx":6C5D
      Tag             =   ":guns"
      Top             =   1440
      Width           =   600
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Index           =   84
      Left            =   2880
      Picture         =   "frmSmileys.frx":706C
      Tag             =   ":gg"
      Top             =   1440
      Width           =   270
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Index           =   83
      Left            =   3240
      Picture         =   "frmSmileys.frx":70F7
      Tag             =   ":tilt"
      Top             =   720
      Width           =   270
   End
   Begin VB.Image imgIcon 
      Height          =   330
      Index           =   82
      Left            =   6120
      Picture         =   "frmSmileys.frx":71D3
      Tag             =   ":guitar"
      Top             =   720
      Width           =   375
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   81
      Left            =   4080
      Picture         =   "frmSmileys.frx":759A
      Tag             =   ":out"
      Top             =   2160
      Width           =   495
   End
   Begin VB.Image imgIcon 
      Height          =   390
      Index           =   80
      Left            =   2280
      Picture         =   "frmSmileys.frx":7695
      Tag             =   ":flush"
      Top             =   1320
      Width           =   390
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   79
      Left            =   2640
      Picture         =   "frmSmileys.frx":7A81
      Tag             =   ":erm"
      Top             =   720
      Width           =   255
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   78
      Left            =   7080
      Picture         =   "frmSmileys.frx":7E08
      Tag             =   ":devil"
      Top             =   120
      Width           =   255
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Index           =   77
      Left            =   4920
      Picture         =   "frmSmileys.frx":7F78
      Tag             =   ":cry"
      Top             =   720
      Width           =   270
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   76
      Left            =   600
      Picture         =   "frmSmileys.frx":8309
      Tag             =   ":bala"
      Top             =   720
      Width           =   255
   End
   Begin VB.Image imgIcon 
      Height          =   345
      Index           =   75
      Left            =   600
      Picture         =   "frmSmileys.frx":83D3
      Tag             =   ":ass"
      Top             =   1320
      Width           =   315
   End
   Begin VB.Image imgIcon 
      Height          =   360
      Index           =   74
      Left            =   960
      Picture         =   "frmSmileys.frx":8484
      Tag             =   ":arnie"
      Top             =   1320
      Width           =   795
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   73
      Left            =   2160
      Picture         =   "frmSmileys.frx":88ED
      Tag             =   ":heart"
      Top             =   720
      Width           =   270
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   72
      Left            =   3720
      Picture         =   "frmSmileys.frx":89B9
      Tag             =   ":wave"
      Top             =   720
      Width           =   375
   End
   Begin VB.Image imgIcon 
      Height          =   375
      Index           =   71
      Left            =   120
      Picture         =   "frmSmileys.frx":8A97
      Tag             =   ":toiletclaw"
      Top             =   2040
      Width           =   705
   End
   Begin VB.Label Label78 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":shoot"
      Height          =   195
      Left            =   5280
      TabIndex        =   28
      Top             =   1800
      Width           =   435
   End
   Begin VB.Label Label77 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":'("
      Height          =   255
      Left            =   6000
      TabIndex        =   27
      Top             =   360
      Width           =   135
   End
   Begin VB.Label Label76 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":baby"
      Height          =   255
      Left            =   5400
      TabIndex        =   26
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label75 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":nono"
      Height          =   255
      Left            =   4800
      TabIndex        =   25
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label74 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":cool"
      Height          =   255
      Left            =   4320
      TabIndex        =   24
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label73 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":smoke"
      Height          =   255
      Left            =   3600
      TabIndex        =   23
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label72 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":?"
      Height          =   255
      Left            =   3240
      TabIndex        =   22
      Top             =   360
      Width           =   135
   End
   Begin VB.Label Label71 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":sleep"
      Height          =   255
      Left            =   2640
      TabIndex        =   21
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label70 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":grr"
      Height          =   255
      Left            =   2160
      TabIndex        =   20
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label69 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":("
      Height          =   255
      Left            =   1680
      TabIndex        =   19
      Top             =   360
      Width           =   135
   End
   Begin VB.Label Label68 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":P"
      Height          =   255
      Left            =   1320
      TabIndex        =   18
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label67 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":D"
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label66 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ";)"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   135
   End
   Begin VB.Label Label65 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":)"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   135
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   70
      Left            =   4920
      Picture         =   "frmSmileys.frx":8ED0
      Tag             =   ":nono"
      Top             =   120
      Width           =   255
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   69
      Left            =   120
      Picture         =   "frmSmileys.frx":904C
      Tag             =   ":)"
      Top             =   120
      Width           =   255
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   68
      Left            =   480
      Picture         =   "frmSmileys.frx":9114
      Tag             =   ";)"
      Top             =   120
      Width           =   255
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   67
      Left            =   840
      Picture         =   "frmSmileys.frx":91DF
      Tag             =   ":D"
      Top             =   120
      Width           =   255
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   66
      Left            =   1680
      Picture         =   "frmSmileys.frx":9570
      Tag             =   ":("
      Top             =   120
      Width           =   255
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   65
      Left            =   2160
      Picture         =   "frmSmileys.frx":9639
      Tag             =   ":grr"
      Top             =   120
      Width           =   255
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   64
      Left            =   2640
      Picture         =   "frmSmileys.frx":99CB
      Tag             =   ":sleep"
      Top             =   120
      Width           =   420
   End
   Begin VB.Image imgIcon 
      Height          =   330
      Index           =   63
      Left            =   3240
      Picture         =   "frmSmileys.frx":9AA4
      Tag             =   ":?"
      Top             =   0
      Width           =   255
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   62
      Left            =   1320
      Picture         =   "frmSmileys.frx":9B7E
      Tag             =   ":P"
      Top             =   120
      Width           =   255
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   61
      Left            =   3720
      Picture         =   "frmSmileys.frx":9F05
      Tag             =   ":smoke"
      Top             =   120
      Width           =   315
   End
   Begin VB.Image imgIcon 
      Height          =   255
      Index           =   60
      Left            =   5400
      Picture         =   "frmSmileys.frx":9FE4
      Tag             =   ":baby"
      Top             =   120
      Width           =   255
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   59
      Left            =   4320
      Picture         =   "frmSmileys.frx":A0C7
      Tag             =   ":cool"
      Top             =   120
      Width           =   255
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   58
      Left            =   6000
      Picture         =   "frmSmileys.frx":A44A
      Tag             =   ":'("
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label64 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":out"
      Height          =   255
      Left            =   4080
      TabIndex        =   17
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label63 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":beat"
      Height          =   255
      Left            =   2880
      TabIndex        =   16
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label62 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":arnie"
      Height          =   195
      Left            =   1200
      TabIndex        =   15
      Top             =   1680
      Width           =   390
   End
   Begin VB.Label Label61 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":cry"
      Height          =   255
      Left            =   4920
      TabIndex        =   14
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label60 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":blah"
      Height          =   195
      Left            =   5400
      TabIndex        =   13
      Top             =   1080
      Width           =   345
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Index           =   57
      Left            =   3360
      Picture         =   "frmSmileys.frx":A7E3
      Tag             =   ":ugly"
      Top             =   1440
      Width           =   240
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   56
      Left            =   5280
      Picture         =   "frmSmileys.frx":A86C
      Tag             =   ":clown"
      Top             =   2160
      Width           =   255
   End
   Begin VB.Image imgIcon 
      Height          =   420
      Index           =   55
      Left            =   1560
      Picture         =   "frmSmileys.frx":AC6F
      Tag             =   ":elk"
      Top             =   1920
      Width           =   450
   End
   Begin VB.Image imgIcon 
      Height          =   270
      Index           =   54
      Left            =   4680
      Picture         =   "frmSmileys.frx":B168
      Tag             =   ":cat"
      Top             =   2160
      Width           =   315
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Index           =   53
      Left            =   1200
      Picture         =   "frmSmileys.frx":B1E9
      Tag             =   ":evil"
      Top             =   720
      Width           =   240
   End
   Begin VB.Image imgIcon 
      Height          =   300
      Index           =   52
      Left            =   2160
      Picture         =   "frmSmileys.frx":B272
      Tag             =   ":drink"
      Top             =   2040
      Width           =   570
   End
   Begin VB.Image imgIcon 
      Height          =   225
      Index           =   51
      Left            =   1680
      Picture         =   "frmSmileys.frx":B419
      Tag             =   ":wow"
      Top             =   720
      Width           =   225
   End
   Begin VB.Image imgIcon 
      Height          =   360
      Index           =   50
      Left            =   6720
      Picture         =   "frmSmileys.frx":B49E
      Tag             =   ":satan"
      Top             =   720
      Width           =   285
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Index           =   49
      Left            =   3600
      Picture         =   "frmSmileys.frx":B5BF
      Tag             =   ":bear"
      Top             =   2160
      Width           =   285
   End
   Begin VB.Label Label59 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":evil"
      Height          =   255
      Left            =   1200
      TabIndex        =   12
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label58 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":bear"
      Height          =   255
      Left            =   3600
      TabIndex        =   11
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label57 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":cat"
      Height          =   255
      Left            =   4680
      TabIndex        =   10
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label56 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":clown"
      Height          =   255
      Left            =   5280
      TabIndex        =   9
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label55 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":elk"
      Height          =   255
      Left            =   1560
      TabIndex        =   8
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label54 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":satan"
      Height          =   195
      Left            =   6720
      TabIndex        =   7
      Top             =   1080
      Width           =   435
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":drink"
      Height          =   255
      Left            =   2040
      TabIndex        =   6
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":wow"
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":ugly"
      Height          =   195
      Left            =   3360
      TabIndex        =   4
      Top             =   1680
      Width           =   330
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   ":uriel"
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   2400
      Width           =   375
   End
   Begin VB.Image imgIcon 
      Height          =   300
      Index           =   48
      Left            =   1080
      Picture         =   "frmSmileys.frx":B647
      Tag             =   ":uriel"
      Top             =   2040
      Width           =   315
   End
End
Attribute VB_Name = "frmSmileys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' uh... icons. Click on them to add them to the end
' of txtSend.text or just check the code...
Private Sub ImgIcon_Click(Index As Integer)
frmMain.txtSend.Text = frmMain.txtSend.Text & imgIcon.Item(Index).Tag
End Sub


Private Sub NChat_Button1_Click()
If FileObj.FolderExists(AppPath & "Smileys") = False Then MkDir (AppPath & "Smileys")
For i = imgIcon.LBound To imgIcon.UBound
SavePicture imgIcon(i).Picture, AppPath & "Smileys\" & "Smiley" & i & ".jpg"
Next i
MsgBox "All smileys have been saved to: " & AppPath & "Smileys", vbInformation, "Done!"

End Sub
