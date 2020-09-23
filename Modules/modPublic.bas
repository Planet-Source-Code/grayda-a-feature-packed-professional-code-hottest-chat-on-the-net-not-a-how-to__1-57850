Attribute VB_Name = "modPublic"
' This form is for our public (and global) calls. Not all of them are here,
' but the rest are categorised

Option Compare Text
' StartMsg and EndMsg are what usernames are enclosed in
' For example: || Grayda || How are you man?
Public StartMSG As String
Public EndMSG As String

' This is what user you picked from frmKick
Public SelUser As String

Public OldTickCount As Long
' Your 'encrypted' administrator password
Public Enc As String

' This string tells you or someone else who your
' last message was from. It is sent as a username,
' and not an IP address
Public LastFrom As String

Public UserName As String

' Allows you to have up to 3 Ini Files open
' Although only one is used in this code :)
Public IniFile(1 To 3) As String

' What files you have shared (*.mp3, *.jpg etc.)
Public FilePattern As String


' This is who the last message was from (in String form
' and not IP form, as the other one is)
Public MFrom As String
' Share you files with the rest of those cheapos? ;)
Public ShareMyFiles As Boolean
' Your Automatic "Away Message"
Public AwayMessage As Boolean
' Whether or not 255.255.255.255 is used or 255.255.255.255
Public Loopback As Boolean
' Popup private messages, rather than hide them
Public Popup As Boolean
' Are you a true admin? (True admins cannot be kicked)
Public TrueAdmin As Boolean
' Allow swearing?
Public Swearing As Boolean
' Show smileys (The pictures, not just the code)?
Public Smiley As Boolean
' Whether or not to show the tip of the day-O
Public DontShowTip As Boolean



' How many NCredits you have
Public NCredits As Long
' How long you have been on NChat for
Public NChatTime As Long
' Number Of Files Shared. What is says
Public NOFS As Long
' Your user icon that you have selected
Public MyIcon As Integer
' How many "icons" you actually own
Public TotalIcons As Integer



Public NewMessage As Boolean

' Allows us to access File System Commands
' Through the Scripting Library Reference
Public FileObj As New Scripting.FileSystemObject

' NChat version. 2 Digit Day, 2 Digit Month,
' 24 hour-hour, and minute
Public Const Ver = "NChat Build10 2010042011"

' Our fancy message setup :)
Public MessageBold As Boolean
Public MessageUnderline As Boolean
Public MessageColour As ColorConstants
Public MessageHColour As ColorConstants


