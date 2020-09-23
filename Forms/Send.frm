VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Send 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sender"
   ClientHeight    =   1320
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar Bar1 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin NChat_Alpha.BinarySender BinarySender1 
      Left            =   3960
      Top             =   480
      _ExtentX        =   794
      _ExtentY        =   794
   End
   Begin VB.Label Label2 
      Caption         =   "000%"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Your file is being transfered to the recipient now. To ensure that they recieve the file, please do not close this window"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "Send"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BinarySender1_CommandAccepted()

Send.BinarySender1.ChunkSize = 4096

Send.BinarySender1.SendFile

End Sub

Private Sub BinarySender1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Text "The file that you tried to send, failed. Please try sending it again. " & Description & vbCrLf, ThatsBad, False
BinarySender1.Reset
BinarySender1.ResetFile

End Sub

Private Sub BinarySender1_SendComplete()
Text File2Send & " has been sent successfully!", svr, True
If frmMain.Visible = False Then Tray.Box File2Send & " has been sent successfully!", "NChat File Transfer"
Sender.ResetFile

End Sub

Private Sub BinarySender1_SendProgress(ByVal Progress As Long, ByVal ProgressMax As Long)
Bar1.Max = ProgressMax
Bar1.Value = Progress
Label2.Caption = Format(Progress, "##") & "%"
End Sub
