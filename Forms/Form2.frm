VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NChat Server Log"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7095
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   8070
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Server Log"
      TabPicture(0)   =   "Form2.frx":2CFA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "RichTextBox1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "sckUDP Log"
      TabPicture(1)   =   "Form2.frx":2D16
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "List2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "sckRooms Log"
      TabPicture(2)   =   "Form2.frx":2D32
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "List1"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   3540
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   6615
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         Height          =   3930
         ItemData        =   "Form2.frx":2D4E
         Left            =   -74880
         List            =   "Form2.frx":2D50
         TabIndex        =   3
         Top             =   480
         Width           =   6615
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   3855
         Left            =   -74880
         TabIndex        =   1
         Top             =   480
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   6800
         _Version        =   393217
         BackColor       =   16777215
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"Form2.frx":2D52
         MouseIcon       =   "Form2.frx":2E26
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Picture Sent: "
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   4080
         Width           =   3975
      End
   End
   Begin MSComDlg.CommonDialog dlgSave 
      Left            =   360
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   4800
      Width           =   6855
   End
   Begin VB.Menu Verb 
      Caption         =   "Verb"
      Visible         =   0   'False
      Begin VB.Menu mnuClearLog 
         Caption         =   "Clear Log"
      End
      Begin VB.Menu mnuSaveLog 
         Caption         =   "Save Log"
      End
   End
   Begin VB.Menu mnuTransmit 
      Caption         =   "Transmit"
      Visible         =   0   'False
      Begin VB.Menu mnuReSend 
         Caption         =   "Re-Send Data"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Edit Data"
      End
      Begin VB.Menu hr1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "Clear Log"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save Log"
      End
   End
End
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This form is some misc stuff to do with logs and broadcast stuff
' Useful for me, and some other people, but not to others

Private Sub Command1_Click()
' Hides the form. If we close it, then we lose all our
' log stuff in the Text box and the list box.
Me.Visible = False

End Sub

Private Sub List2_DblClick()
' Re-sends the data we are clicking on.
' Good for re-sending a message, kick request etc.
mnuReSend_Click
End Sub

Private Sub List2_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 2 Then
PopupMenu mnuTransmit
End If
End Sub

Private Sub mnuClear_Click()
List2.Clear

End Sub

Private Sub mnuClearLog_Click()
RichTextBox1.Text = ""

End Sub

Private Sub mnuEdit_Click()
Dim Tmp As String

Tmp = InputBox("Enter (Or Edit) the data that you wish to send and Press OK", "Edit Data", List2.List(List2.ListIndex))
If Tmp > "" Then Broadcast Tmp
End Sub

Private Sub mnuReSend_Click()
If List2.Text <> "" Then Broadcast List2.Text
End Sub

Private Sub mnuSave_Click()
' Save the contents of List2
dlgSave.Filter = "Text files (*.txt)|*.txt"
dlgSave.ShowSave
Open dlgSave.filename For Output As #1

For i = 0 To List2.ListCount
Print #1, List2.List(i)
Next i
Close #1
End Sub

Private Sub mnuSaveLog_Click()
' Saves log data
' If the selected type is *.txt or *.*
' then save it as plain text.
' If it's not, then save it with RTF tags etc.
dlgSave.Filter = "Text Files (*.TXT)|*.txt|RTF File (*.rtf)|*.RTF|All Files (*.*)|*.*"
dlgSave.ShowSave
If dlgSave.filename = "" Then Exit Sub
Open dlgSave.filename For Output As #1
If dlgSave.FilterIndex = 1 Or dlgSave.FilterIndex = 3 Then
Print #1, RichTextBox1.Text
ElseIf dlgSave.FilterIndex = 2 Then
Print #1, RichTextBox1.TextRTF
End If
Close #1
Text "Log Saved!" & vbCrLf, svr, True
End Sub
Private Sub RichTextBox1_Change()
' Not selecting anything? Then scroll to the end
If RichTextBox1.SelLength = 0 Then RichTextBox1.SelStart = Len(RichTextBox1.Text)
End Sub


