VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl BinarySender 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   InvisibleAtRuntime=   -1  'True
   Picture         =   "BinarySender.ctx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "BinarySender.ctx":0C42
   Begin VB.Timer tmrUploadSpeed 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   0
      Top             =   480
   End
   Begin MSWinsockLib.Winsock wsSender 
      Left            =   2040
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   3000
   End
   Begin MSWinsockLib.Winsock wsInfo 
      Left            =   1560
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   1700
   End
End
Attribute VB_Name = "BinarySender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mCurrentFileSize As Long
Dim mCurrentFileName As String

Public ChunkSize As Long

Dim TotalSent As Long

Dim SourceFilename As String

Dim ByteNow As Long

Dim t As Integer

Dim UploadSpeed As Long
Dim UploadSecond As Long

'Events
Public Event Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Public Event SendError(ByVal Number As Long, Description As String)
Public Event Connect()
Public Event CommandAccepted()
Public Event CommandRefused()
Public Event SendProgress(ByVal Progress As Long, ByVal ProgressMax As Long)
Public Event SendComplete()
Public Event SpeedRecord(ByVal Speed As Long)

Public Property Get CurrentFileSize() As Long
CurrentFileSize = mCurrentFileSize
End Property

Public Property Get CurrentFileName() As String
CurrentFileName = mCurrentFileName
End Property

Public Sub ResetFile()

mCurrentFileSize = 0
mCurrentFileName = ""

ChunkSize = 4096

TotalSent = 0

SourceFilename = ""

ByteNow = 0

UploadSpeed = 0

UploadSecond = 0

tmrUploadSpeed.Enabled = False

End Sub

Public Sub Reset()

mCurrentFileSize = 0
mCurrentFileName = ""

ChunkSize = 4096

TotalSent = 0

SourceFilename = ""

ByteNow = 0

UploadSpeed = 0

UploadSecond = 0

wsInfo.Close

wsSender.Close

tmrUploadSpeed.Enabled = False

End Sub

Private Sub tmrUploadSpeed_Timer()

UploadSpeed = TotalSent - UploadSecond

RaiseEvent SpeedRecord((UploadSpeed / 1024) * 2)

UploadSecond = TotalSent

End Sub

Private Sub UserControl_Resize()
UserControl.Width = 450
UserControl.Height = 450
End Sub

Private Sub UserControl_InitProperties()
ChunkSize = 4096
RemotePortBinary = 3000
RemotePortInfo = 1700
End Sub

Public Property Get TheWinsock() As Winsock

Set TheWinsock = wsSender

End Property

Public Property Get RemoteHost() As String
RemoteHost = wsSender.RemoteHost
End Property

Public Property Let RemoteHost(Host As String)
wsSender.RemoteHost = Host
End Property

Public Property Get RemoteHostIP() As String
RemoteHostIP = wsSender.RemoteHostIP
End Property

Public Property Let RemotePortBinary(Port As Long)
wsSender.RemotePort = Port
End Property

Public Property Get RemotePortBinary() As Long
RemotePortBinary = wsSender.RemotePort
End Property

Public Property Let RemotePortInfo(Port As Long)
wsInfo.RemotePort = Port
End Property

Public Property Get RemotePortInfo() As Long
RemotePortInfo = wsInfo.RemotePort
End Property

Public Property Let Source(str As String)
On Error Resume Next
SourceFilename = str
mCurrentFileName = StripPath(str)
mCurrentFileSize = FileLen(str)
End Property

Public Property Get Source() As String
Source = SourceFilename
End Property

Public Sub Connect()

    With wsInfo
        .Close
        .RemoteHost = Me.RemoteHost
        .RemotePort = Me.RemotePortInfo
        DoEvents
        .Connect
    End With
    
End Sub

Public Sub SendInfo()
On Error Resume Next
wsInfo.SendData "FIS" & mCurrentFileSize & "|" & 0 & "@" & mCurrentFileName
End Sub

Private Sub wsInfo_DataArrival(ByVal bytesTotal As Long)

Dim a As String
wsInfo.GetData a

Select Case Left(a, 3)

    Case "RFC"
        tmrUploadSpeed.Enabled = False
        RaiseEvent SendComplete
    
    Case "RFS"
        RaiseEvent CommandRefused
    
    Case "ACP"
        RaiseEvent CommandAccepted
        
    Case "CNT"
        RaiseEvent Connect
        
        With wsSender
            .Close
            .RemoteHost = Me.RemoteHost
            .RemotePort = Me.RemotePortBinary
            DoEvents
            .Connect
        End With

End Select

End Sub

Public Sub SendFile()
On Error Resume Next
    Dim bytBuf() As Byte
    t = FreeFile
        
        Dim i As Long
        
        tmrUploadSpeed.Enabled = True
        If SourceFilename = "" Then Exit Sub
        Open SourceFilename For Binary Access Read As #t

            ReDim bytBuf(1 To ChunkSize) As Byte
        
            Do Until (CurrentFileSize - ByteNow) < ChunkSize
                    
                    DoEvents
                    Get #t, ByteNow + 1, bytBuf()
                    
                    ByteNow = ByteNow + ChunkSize
                    
                    DoEvents
                    On Error GoTo SendError
                    wsSender.SendData bytBuf
            
            Loop
            
            Dim LastChunkSize As Long
            LastChunkSize = CurrentFileSize - ByteNow
            
            DoEvents
            ReDim bytBuf(1 To LastChunkSize) As Byte
            Get #t, ByteNow + 1, bytBuf()
            
            ByteNow = ByteNow + LastChunkSize
            
            DoEvents
            wsSender.SendData bytBuf
            
            Close #t
            
        tmrUploadSpeed.Enabled = False

Exit Sub
SendError:
RaiseEvent SendError(Err.Number, Err.Description)
tmrUploadSpeed.Enabled = False
End Sub

Private Sub wsInfo_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
RaiseEvent Error(Number, Description, Scode, Source, HelpFile, HelpContext, CancelDisplay)
End Sub

Private Sub wsSender_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
RaiseEvent Error(Number, Description, Scode, Source, HelpFile, HelpContext, CancelDisplay)
End Sub

Private Sub wsSender_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
TotalSent = TotalSent + bytesSent

    DoEvents
    RaiseEvent SendProgress(TotalSent, mCurrentFileSize)

End Sub
