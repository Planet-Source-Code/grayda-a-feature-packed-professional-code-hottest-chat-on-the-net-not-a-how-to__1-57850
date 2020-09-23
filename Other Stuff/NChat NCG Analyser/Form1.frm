VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "NChat Data Decrypter v1.0"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   5970
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   4560
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog dlgLoad 
      Left            =   1680
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "NChat Settings File (*.ncg)|*.ncg"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Analyse"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4560
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   4320
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Gets the proper windows username, instead of using environ("username")
Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long

Private Function GetUserName() As String
' Simple sub to get our windows username
   Dim UserName2 As String * 255
   Call GetUserNameA(UserName2, 255)
   GetUserName = Left$(UserName2, InStr(UserName2, Chr$(0)) - 1)
End Function

Private Sub LoadSettings()
' Loads your NCredits, Username etc.
' Most things need to be decoded before they are loaded
On Error Resume Next
Dim TempStr As String

' Pretty obvious: Opens your file and reads from it
Open dlgLoad.FileName For Input As #1
' Whether or not your files are shared
Line Input #1, TempStr
List1.AddItem "Files Shared?: " & TempStr

Line Input #1, TempStr2
List1.AddItem "Shared File Path: " & TempStr2
Line Input #1, TempStr
' What files you have shared (Like *.mp3;*.jpg etc.)
List1.AddItem "Type of file shared: " & TempStr
' Some simple user info stuff. Read the variables
' to find out what is loaded
Line Input #1, TempStr

List1.AddItem "Username: " & Decode(TempStr, GetUserName & "ŽZ5¬")
' No username set? Then it's set as your windows username

' Number of user icons you have "Purchased" from the NChat
' shop (THere are like 50 of them to buy!!)
Line Input #1, TempStr
List1.AddItem "Number of Icons available: " & Decode(TempStr, GetUserName & "ŽZ5¬")

' Whether or not you are an admin
Line Input #1, TempStr
Dim Ad As Boolean
If Decode(TempStr, GetUserName & "ŽZ5¬") = "TharSheBlows" Then

Ad = False
Else
Ad = True
End If

List1.AddItem "Administrator?: " & Ad

Line Input #1, TempStr
List1.AddItem "NCredits: " & Decode(TempStr, GetUserName & "ŽZ5¬")

Line Input #1, TempStr
List1.AddItem "Username Check: " & Decode(TempStr, GetUserName & "ŽZ5¬")
DoEvents

Line Input #1, TempStr
' Your icon next to your username
List1.AddItem "Icon #: " & TempStr
Line Input #1, TempStr
' Is swearing filtered out?
List1.AddItem "Swearing?: " & TempStr
Line Input #1, TempStr
'Line Input #1, TempStr
' Your last Profile (INI) Loaded
List1.AddItem "Last Profile: " & Decode(TempStr, GetUserName & "ŽZ5¬")
Line Input #1, TempStr
' See top of this form code for trueadmin info
List1.AddItem "True Admin?: " & Decode(TempStr, GetUserName & "ŽZ5¬")
' Popup Private Messages (Rather than hide them)?
Line Input #1, TempStr
List1.AddItem "Popup PMs?: " & TempStr
Line Input #1, TempStr
List1.AddItem "Show Tip of the day?: " & TempStr

Line Input #1, TempStr
List1.AddItem "Message Bold?: " & Decode(TempStr, GetUserName & "©®©32")
Line Input #1, TempStr

List1.AddItem "Message Underlined?: " & Decode(TempStr, GetUserName & "©®©32")
Line Input #1, TempStr
List1.AddItem "Message Colour: " & Decode(TempStr, GetUserName & "©®©32")
Line Input #1, TempStr
List1.AddItem "Message Highlight Colour: " & Decode(TempStr, GetUserName & "©®©32")
Line Input #1, TempStr
' The start of the message (eg. || Username || Hello!)
List1.AddItem "Start Message: " & TempStr
' Are you permanently banned from NChat?
Line Input #1, TempStr
' The end of the message, like Startmsg
List1.AddItem "End Message: " & TempStr

Line Input #1, TempStr
' Are you permanently banned from NChat?
Tmp = Decode(TempStr, "12345678910")
If Tmp = "Banned" Then
MsgBox "YOU HAVE BEEN BANNED FROM NCHAT. NCHAT WILL NOW CLOSE", vbCritical, "BANNED"
End
End If


Close #1
End Sub

Private Sub Command1_Click()


dlgLoad.ShowOpen
LoadSettings
End Sub

Private Sub Command2_Click()
If dlgLoad.FileName = "" Then
MsgBox "Please load a file first!!", vbCritical, "No file loaded"
Exit Sub
End If
List1.Clear
LoadSettings
End Sub
