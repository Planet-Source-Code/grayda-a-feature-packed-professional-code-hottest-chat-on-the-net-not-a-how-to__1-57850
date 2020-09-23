Attribute VB_Name = "modWSock"
Option Compare Text

' Is NChat running on a LAN (ie. Connectionless UDP)
' or is it running over the internet (ie. TCP/IP Connection)
Public IsInternet As Boolean

' Data recieved through sckUDP (in frmMain
Public sData As String
' Results of the SplitVB5 function
Public Result() As String
' The sckUDP and sckRooms addresses that are to be used
' They are string because of the decimal points inbetween
Public Address As String
' For file Transfer
Public File2Send As String
Public File2Save As String

Public Sub Broadcast(CData As String)
On Error Resume Next
' Sends data to EVERYONE in the room

' These are text shortcuts. When you type these,
' they are replaced with correct values. this makes
' chatting easier, because Welcome Messages, Notch
' and other stuff can be more dynamic
CData = Replace(CData, "+username+", UserName)
CData = Replace(CData, "+ip+", frmMain.sckUDP.LocalIP)
CData = Replace(CData, "+room+", RoomName & " (" & Room & ")")
CData = Replace(CData, "+ncredits+", NCredits)
CData = Replace(CData, "+newuser+", NewUser)
CData = Replace(CData, "+ver+", Ver)
CData = Replace(CData, "+ntime+", NChatTime)
CData = Replace(CData, "+time+", Format(Time, "HH:mm"))
' +d+ is our delimiter for sent data. This is simpler
' than typing Alt+0248
CData = Replace(CData, "+d+", "ø")
' Who our last message is from
CData = Replace(CData, "+lastfrom+", LastFrom)
CData = Replace(CData, "+roomhost+", RoomHost)
Randomize
CData = Replace(CData, "+someguy+", frmMain.List1.ListItems.Item(Int(Rnd * frmMain.List1.ListItems.Count) + 1).Text)

FindPos = InStr(1, CData, "+result")
If FindPos > 0 Then CData = Replace(CData, "+result" & Mid(CData, FindPos + Len("+result"), 1) & "+", Result(Mid(CData, FindPos + Len("+result"), 1)))

If IsInternet = False Then
frmMain.sckUDP.Close
' 'Resets' the connection
frmMain.sckUDP.LocalPort = frmMain.sckUDP.RemotePort
frmMain.sckUDP.RemoteHost = Address
frmMain.sckUDP.RemotePort = Room

frmMain.sckUDP.Connect
End If

' Finally sends the data
' NChat©°® is our data 'header'.
' If the incoming data is missing this, then
' ignore it
frmMain.sckUDP.SendData "N©H@-|-ø" & CData
If Left(CData, 4) = "Move" Or Left(CData, 3) = "usr" Or Left(CData, 5) = "Clear" Or Left(CData, 6) = "Colour" Or Left(CData, 4) = "Size" Then Exit Sub
If Trim(CData) > "" Then
frmLog.List2.AddItem CData
Else
frmLog.List2.AddItem "---Blank Data!---"
End If


End Sub

Public Sub RoomBroadcast(BText As String)
' This is to do with the list of rooms
' See the sckRooms_Dataarival sub for more info

' 'Resets' the connection
On Error Resume Next
If IsInternet = False Then
frmMain.sckRooms.Close
frmMain.sckRooms.RemoteHost = Address
frmMain.sckRooms.LocalPort = 127
frmMain.sckRooms.RemotePort = 127
End If
'frmMain.sckRooms.Connect
' Finally sends the data
frmMain.sckRooms.SendData BText

End Sub

' uuhhh... I think these 3 subs are here for the
' binary transfer and Receiver usercontrols
Sub SaveBinaryArray(ByVal filename As String, WriteData() As Byte)

    Dim T As Integer
    T = FreeFile
    Open filename For Binary Access Write As #T
        
            Put #T, , WriteData()
        
    Close #T
    
End Sub

Function ReadBinaryArray(ByVal Source As String)

    Dim bytBuf() As Byte
    
    Dim T As Integer
    T = FreeFile
    
    Open Source For Binary Access Read As #T
    
    Dim n As Long
    
    ReDim bytBuf(1 To LOF(T)) As Byte
    Get #T, , bytBuf()
    
    ReadBinaryArray = bytBuf()
    
    Close #T
    
End Function

Public Function StripPath(T As String) As String

  Dim X As Integer
  Dim ct As Integer

    StripPath = T
    X = InStr(T, "\")
    Do While X
        ct = X
        X = InStr(ct + 1, T, "\")
    Loop
    If ct > 0 Then StripPath = Mid$(T, ct + 1)

End Function


