VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl BinaryReceiver 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4965
   InvisibleAtRuntime=   -1  'True
   Picture         =   "BinaryReceiver.ctx":0000
   ScaleHeight     =   3360
   ScaleWidth      =   4965
   ToolboxBitmap   =   "BinaryReceiver.ctx":0C42
   Begin VB.Timer tmrDownloadSpeed 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   0
      Top             =   480
   End
   Begin MSWinsockLib.Winsock wsReader 
      Left            =   1920
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   3000
   End
   Begin MSWinsockLib.Winsock wsInfo 
      Left            =   1440
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   1700
   End
End
Attribute VB_Name = "BinaryReceiver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim TotalByteNow As Long
Dim ByteNow As Long

Dim mCurrentFileSize As Long
Dim mCurrentFileName As String

Dim PackageCount As Long 'Just Dummy

Dim TargetFilename As String
Dim WritePos As Long

Dim DownloadSecond As Long
Dim DownloadSpeed As Long

Dim DataIN() As Byte
Dim t As Integer

Dim mSaveTarget As String
'Dim PCount As Long

Public Event Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Public Event SendRequest()
Public Event ConnectionRequest()
Public Event ReceiveComplete()
Public Event ReceiveProgress(ByVal Progress As Long, ByVal ProgressMax As Long)
Public Event SpeedRecord(ByVal Speed As Long)

Public Sub ResetFile()

TotalByteNow = 0
ByteNow = 0

mCurrentFileSize = 0
mCurrentFileName = ""

TargetFilename = ""
WritePos = 0

DownloadSecond = 0
DownloadSpeed = 0

t = FreeFile

tmrDownloadSpeed.Enabled = False

End Sub

Public Sub Reset()

TotalByteNow = 0
ByteNow = 0

mCurrentFileSize = 0
mCurrentFileName = ""

TargetFilename = ""
WritePos = 0

wsInfo.Close
wsReader.Close

DownloadSecond = 0
DownloadSpeed = 0

t = FreeFile

tmrDownloadSpeed.Enabled = False

End Sub

Public Property Get CurrentFileSize() As Long
CurrentFileSize = mCurrentFileSize
End Property

Public Property Get CurrentFileName() As String
CurrentFileName = mCurrentFileName
End Property

Private Sub tmrDownloadSpeed_Timer()

DownloadSpeed = TotalByteNow - DownloadSecond

RaiseEvent SpeedRecord((DownloadSpeed / 1024) * 2)

DownloadSecond = TotalByteNow

End Sub

Private Sub UserControl_InitProperties()
LocalPortBinary = 3000
LocalPortInfo = 1700
End Sub

Private Sub UserControl_Resize()
UserControl.Width = 450
UserControl.Height = 450
End Sub

Public Property Get TheWinsock() As Winsock

Set TheWinsock = wsReader

End Property

Public Property Get LocalHostName() As String
LocalHostName = wsReader.LocalHostName
End Property

Public Property Get LocalIP() As String
LocalIP = wsReader.LocalIP
End Property

Public Property Let LocalPortBinary(Port As Long)
wsReader.LocalPort = Port
End Property

Public Property Get LocalPortBinary() As Long
LocalPortBinary = wsReader.LocalPort
End Property

Public Property Let LocalPortInfo(Port As Long)
wsInfo.LocalPort = Port
End Property

Public Property Get LocalPortInfo() As Long
LocalPortInfo = wsInfo.LocalPort
End Property

Private Sub wsInfo_ConnectionRequest(ByVal requestID As Long)
wsInfo.Close
wsInfo.Accept requestID

DoEvents
wsInfo.SendData "CNT"

End Sub

Private Sub wsInfo_DataArrival(ByVal bytesTotal As Long)
Dim a As String
wsInfo.GetData a

Select Case Left(a, 3)
    
    Case "FSC"
        RaiseEvent ReceiveComplete
    
    Case "FIS"
        mCurrentFileSize = CLng(Mid(a, 4, InStr(1, a, "|") - 4))
        PackageCount = Mid(a, 4 + Len(Trim(str(CurrentFileSize))) + 1, InStr(1, a, "@") - (4 + Len(Trim(str(CurrentFileSize)))) - 1)
        mCurrentFileName = Mid(a, 5 + Len(Trim(str(CurrentFileSize))) + Len(str(Trim(PackageCount))))
        
        RaiseEvent SendRequest
        
End Select

End Sub

Private Sub wsInfo_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
RaiseEvent Error(Number, Description, Scode, Source, HelpFile, HelpContext, CancelDisplay)
End Sub

Private Sub wsReader_ConnectionRequest(ByVal requestID As Long)
wsReader.Close
wsReader.Accept requestID
RaiseEvent ConnectionRequest
End Sub

Public Sub AcceptSendRequest(SaveTarget As String)

On Error Resume Next

t = FreeFile

mSaveTarget = SaveTarget

Open mSaveTarget & "TMP" For Binary Access Write As #t

DoEvents
wsInfo.SendData "ACP"

End Sub

Public Sub RefuseSendRequest()

DoEvents
wsInfo.SendData "RFS"

End Sub

Private Sub wsReader_DataArrival(ByVal bytesTotal As Long)

On Error Resume Next

ReDim DataIN(1 To bytesTotal) As Byte

'DoEvents
wsReader.GetData DataIN()

tmrDownloadSpeed.Enabled = True

'PCount = PCount + 1

'DoEvents
        Put #t, WritePos + 1, DataIN()
        WritePos = WritePos + bytesTotal
        TotalByteNow = TotalByteNow + bytesTotal
        
'If PCount Mod 20 = 0 Then
    DoEvents
    RaiseEvent ReceiveProgress(TotalByteNow, mCurrentFileSize)
'End If

If TotalByteNow >= CurrentFileSize Then
'If LOF(t) >= CurrentFileSize Then

        Do Until LOF(t) >= TotalByteNow Or LOF(t) >= CurrentFileSize
        Loop
        'Receive Complete
        Close #t
        
        FileCopy mSaveTarget & "TMP", mSaveTarget
        Kill mSaveTarget & "TMP"
        
        tmrDownloadSpeed.Enabled = False
        RaiseEvent ReceiveComplete
        FileSize = 0
                
        DoEvents
        wsInfo.SendData "RFC"
        
End If

End Sub

Private Sub wsReader_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
RaiseEvent Error(Number, Description, Scode, Source, HelpFile, HelpContext, CancelDisplay)
End Sub

Public Sub Listen()

With wsInfo
    .Close
    .LocalPort = Me.LocalPortInfo
    .Listen
End With

With wsReader
    .Close
    .LocalPort = Me.LocalPortBinary
    .Listen
End With

End Sub
