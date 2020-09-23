VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Receiver 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reciever"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin NChat_Alpha.BinaryReceiver BinaryReceiver1 
      Left            =   4200
      Top             =   0
      _ExtentX        =   794
      _ExtentY        =   794
   End
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "000%"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "The file you have chosen, is being downloaded now. Please do not close this window until the bar has reached 100%"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "Receiver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BinaryReceiver1_ConnectionRequest()
BinaryReceiver1.AcceptSendRequest File2Save
End Sub

Private Sub BinaryReceiver1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
MsgBox "File not downloaded!!" & vbCrLf & vbCrLf & Number & vbCrLf & Description, vbCritical, "File not downloaded"
BinaryReceiver1.Reset
BinaryReceiver1.ResetFile
BinaryReceiver1.Listen

End Sub

Private Sub BinaryReceiver1_ReceiveComplete()
MsgBox "File Downloaded!!", vbInformation, "DONE!!"
If Me.Visible = False Then Tray.Box File2Save & " has been downloaded successfully!", "NChat File Transfer"
BinaryReceiver1.Reset
BinaryReceiver1.ResetFile
BinaryReceiver1.Listen

End Sub

Private Sub BinaryReceiver1_ReceiveProgress(ByVal Progress As Long, ByVal ProgressMax As Long)
On Error Resume Next
Bar1.Max = ProgressMax
Bar1.Value = Progress
Label2.Caption = Format(Progress, "###") & "%"
End Sub

Private Sub BinaryReceiver1_SendRequest()
BinaryReceiver1.AcceptSendRequest File2Save
End Sub
