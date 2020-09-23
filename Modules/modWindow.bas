Attribute VB_Name = "modWindow"
' Say you are talking to user2 and user3,
' Any message from user2 will only come into user2's
' window. Messages from user3 will go into it's
' respective window. If user4 joins in, a new window
' will be created for them
Public CW() As New frmChat

' Consts for Always On Top stuff
Public Const HWND_TOPMOST = -1
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40

' This fades out the window when you exit NChat
' API calls and Constants from allapi.net's API List
Public Const AW_HIDE = &H10000 'Hides the window. By default, the window is shown.
Public Const AW_BLEND = &H80000 'Uses a fade effect. This flag can be used only if hwnd is a top-level window.
Public Declare Function AnimateWindow Lib "user32" (ByVal hwnd As Long, ByVal dwTime As Long, ByVal dwFlags As Long) As Boolean

' Sets the position of the window (Including Z Order)
' This code from allapi.net's API Guide
Public Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

' All these colour constants are for profiles
' They pretty much describe what they colour

' Heading: Headings and large text
' Msg: Standard user messages
' Dis: User Disconnects
' Svr: Server (Purple) Messages
' Act: WinMX Actions (Emotes)
' Con: User connects
' ThatsGood: Good news
' ThatsBad: Bad news / Errors
Public Heading As ColorConstants
Public Msg As ColorConstants
Public dis As ColorConstants
Public svr As ColorConstants
Public act As ColorConstants
Public con As ColorConstants
Public ThatsGood As ColorConstants
Public ThatsBad As ColorConstants
Public Col As ColorConstants

' Caption prefix is the form's message. For example,
' If Grayda sends the message Hi, and my
' CaptionPrefix is: NChat [
' EndWindow is: ]
' Then my form caption will be:
' NChat [ Grayda - Hi ]
Public CaptionPrefix As String
Public EndWindow As String

' OldCap is the first part of the frmMain Caption
' (eg. Welcome to NChat - <Rest of the caption goes here>)
Public OldCap As String

' System tray stuff
Public Tray As New cSysTray

Public Sub OnTop(window As Long)
    'Set the window position to topmost
    SetWindowPos window, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    'KPD-Team 1998
    'URL: http://www.allapi.net/
    'E-Mail: KPDTeam@Allapi.net
End Sub

Public Sub Status(Tip As String, Optional Icon As Long)
' The Status sub does many things:
' It sets frmMain's caption, sets the Tray's tip
' and if necessary, sets the tray's icon
OldCap = frmMain.Caption

If Tip > "" Then
frmMain.Caption = Tip
Tray.TText = Tip
Tray.Update
End If

If Icon > 0 Then
Tray.IconHandle = Icon
End If

End Sub

Public Function ShowBox(ButtonText As String, WindowTitle As String) As String
' Showbox lets us do many things.
' for example, it lets us pick one username from the
' list of users, so we can kick them, print their info
' or anything that requires the end user to pick a name

frmKick.Show
frmKick.Command1.Caption = ButtonText
frmKick.Caption = WindowTitle
Do Until frmKick.Visible = False
DoEvents
DoEvents
Loop
ShowBox = SelUser
End Function
