Attribute VB_Name = "modRTB"
Option Compare Text

' These consts and types
' control many things,
' but for now, we will use them
' to highlight stuff (like with a
' flourecent highlighted texta marker thing)
Public Const WM_USER = &H400
Public Const SCF_SELECTION = &H1&
Public Const EM_SETCHARFORMAT = (WM_USER + 68)
Public Const CFM_BACKCOLOR = &H4000000

Public Const LF_FACESIZE = 32
Public Type CHARFORMAT2
    cbSize As Integer
    wPad1 As Integer
    dwMask As Long
    dwEffects As Long
    yHeight As Long
    yOffset As Long
    crTextColor As Long
    bCharSet As Byte
    bPitchAndFamily As Byte
    szFaceName(0 To LF_FACESIZE - 1) As Byte
    wPad2 As Integer
    wWeight As Integer
    sSpacing As Integer
    crBackColor As Long
    lLCID As Long
    dwReserved As Long
    sStyle As Integer
    wKerning As Integer
    bUnderlineType As Byte
    bAnimation As Byte ' ooooh...... I don't know what this does :(
    bRevAuthor As Byte
    bReserved1 As Byte
End Type

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lparam As Any) As Long
Public Declare Function SendMessageByVal Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lparam As Long) As Long

Public Sub HighlightText(Colour As ColorConstants)
Dim udtCharFormat As CHARFORMAT2
        udtCharFormat.dwMask = CFM_BACKCOLOR
        udtCharFormat.cbSize = LenB(udtCharFormat)
        udtCharFormat.crBackColor = Colour
Call SendMessageByVal(frmMain.txtChat.hwnd, EM_SETCHARFORMAT, SCF_SELECTION, VarPtr(udtCharFormat))


End Sub

Public Sub Text(Text As String, Optional Colour As ColorConstants, Optional Bold As Boolean, Optional Italic As Boolean, Optional Underline As Boolean, Optional Size As Integer, Optional Alignment As AlignmentConstants, Optional Font As String, Optional CheckSmileys As String)
On Error Resume Next
' Puts text into the rich text box
If Text = "" Then Exit Sub
If CheckSmileys = "" Then CheckSmileys = "True"
' Replace shortcuts with their values
Text = Replace(Text, "+username+", UserName)
Text = Replace(Text, "+ip+", frmMain.sckUDP.LocalIP)
Text = Replace(Text, "+room+", RoomName & " (" & Room & ")")
Text = Replace(Text, "+ncredits+", NCredits)
Text = Replace(Text, "+newline+", vbCrLf)
Text = Replace(Text, "+newuser+", NewUser)
Text = Replace(Text, "+remoteip+", frmMain.sckUDP.RemoteHostIP)
Text = Replace(Text, "+ntime+", NChatTime)
Text = Replace(Text, "+ver+", Ver)
Text = Replace(Text, "+time+", Format(Time, "HH:mm"))

' Room Stuff
Text = Replace(Text, "+roomname+", RoomName)
Text = Replace(Text, "+roomid+", Room)
Text = Replace(Text, "+roomuser+", frmMain.List1.ListItems.Count)
Text = Replace(Text, "+roomhost+", RoomHost)

With frmMain.txtChat
' Set the cursor at the end
    .SelStart = Len(.Text)
' The length of the selection should be 0
    .SelLength = Len(.Text)
    
    .SelBold = Bold
    If Font > "" Then .SelFontName = Font
    .SelItalic = Italic
    .SelUnderline = Underline
    .SelFontSize = Size
    .SelAlignment = Alignment
    .SelColor = Colour
    
    .SelText = Text
    .SelStart = Len(.Text)
    .SelLength = 0

End With
If CheckSmileys = "True" Then CheckRTF frmMain.txtChat
End Sub

Public Sub CheckRTF(CheckOBJ As Object)

' THE code to turn the :) and :( etc. into pictures
Dim StartAt As Long
'On Error Resume Next
Dim TheSmiley As String
' Only print the smiley once?
Dim DoOnce As Boolean

' Allows you to have (some) html tags in text
With CheckOBJ

DoOnce = True
If Smiley = True Then



t = 47
' All the smileys in frmSmileys
For t = frmSmileys.imgIcon.LBound To frmSmileys.imgIcon.UBound
' The smiley to check for

TheSmiley = frmSmileys.imgIcon(t).Tag
' A quicker way to determine where the smileys start
StartAt = InStr(1, .Text, TheSmiley, vbBinaryCompare) - 1
If StartAt > 0 Then
' Gets rid of the code :) :devil :poo etc.
.SelStart = StartAt
.SelLength = Len(frmSmileys.imgIcon(t).Tag)
.SelText = Replace(.SelText, frmSmileys.imgIcon(t).Tag, "")
' Unlocks the txtChat control, so we can
' paste our smiley into it
'.Locked = False
.SelStart = StartAt
' Paste the damn thing!! :)

AddPic frmSmileys.imgIcon(t).Picture, CheckOBJ

' Lock the control off again
'[.Locked = True
DoOnce = False

' If there is more than one smiley on the line,
' then don't put in a vbcrlf for each one!!
If DoOnce = True Then
Text "" & vbCrLf
DoOnce = False
End If

End If

Next t
'Next b

End If
End With
End Sub

Public Sub Txt2(Text As String, Colour As ColorConstants, ByRef window As Integer)
' Writes text into a private chat window, hence
' the extra 'window' syntax
Text = Replace(Text, "+username+", UserName)
Text = Replace(Text, "+ip+", frmMain.sckUDP.LocalIP)
Text = Replace(Text, "+room+", RoomName)
Text = Replace(Text, "+ncredits+", NCredits)

With CW(window).Text1
    .SelStart = Len(.Text)
    .SelLength = Len(.Text)
    .SelColor = Colour
    .SelText = Text
    .SelLength = 0
End With

End Sub

Public Sub Log(Text As String, Optional Colour As ColorConstants, Optional Bold As Boolean)
' Custom Text commands
' Like the Text Sub in module1, but much smaller

' Writes some stuff into our administrator's log
' (frmLog).

With frmLog.RichTextBox1
    .SelStart = Len(.Text)
    .SelLength = Len(.Text)
    .SelColor = Colour
    .SelBold = Bold
    .SelText = Text
    .SelLength = 0

End With

End Sub

