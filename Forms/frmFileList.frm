VERSION 5.00
Begin VB.Form frmFileList 
   Caption         =   "Files you have shared"
   ClientHeight    =   3030
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4620
   Icon            =   "frmFileList.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4620
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   2760
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmFileList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This form allows you to see what files you have shared
' This is also where your list of files is sent to others from
Private Sub Form_Resize()
File1.Move 150, 150, Me.Width - 450, Me.Height - 750

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
frmFileList.File1.Pattern = FilePattern
End Sub
