Attribute VB_Name = "modImages"
Option Compare Text
' OK, let me make some things clear:

' 1) I didn't write ANY of this code in this module
' 2) UDMX IOCP© wrote the WMF stuff, and is used
'    with permission (go to www.intra-dream.com to
'    see more of his, and other's work)
' 3) An unknown author wrote the GDI PNG stuff,
'    PLEASE go to
'    http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=55313&lngWId=1
'    to see the proper example. AFAIK, I can use this
'    code, and the author would appreciate it if
'    you voted for THE code, as I am going to do
'    5 globes!!

' Stuff for making things transparent :)
Private Const WS_EX_TRANSPARENT = &H20&
Private Const GWL_EXSTYLE = (-20)

Public ThePic As String
Public PicStarted As Boolean

' This is for both the Per-Pixel PNG rendering,
' and
Public Declare Function SetWindowLong Lib "user32" _
  Alias "SetWindowLongA" (ByVal hwnd As Long, _
  ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

' Our API to flood fill a picture box
Public Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long

' Just some simple Publics for the whiteboard
' They tell you where your last line was drawn from
Public lastX As Long
Public lastY As Long

'This was from VBElite
'It simply creates the RTB Syntax for any pictures
'When you paste an image... then select it [highlight]
'And check the SElRTF you'll see the Syntax for that picture
'This makes that syntax from just a picture.. STDPIC


Private Type Size
    cx As Long
    cy As Long
End Type


Private Type bitmap
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

'Private Type METAHEADER
'    mtType As Integer
'    mtHeaderSize As Integer
'    mtVersion As Integer
'    mtSize As Long
'    mtNoObjects As Integer
'    mtMaxRecord As Long
'    mtNoParameters As Integer
'End Type

' Used to create the metafile
Private Declare Function CreateMetaFile Lib "gdi32" Alias "CreateMetaFileA" (ByVal lpString As String) As Long
Private Declare Function CloseMetaFile Lib "gdi32" (ByVal hDCMF As Long) As Long
Private Declare Function DeleteMetaFile Lib "gdi32" (ByVal hMF As Long) As Long
' 6 APIs used to render/embed the bitmap in the metafile
Private Declare Function SetMapMode Lib "gdi32" (ByVal hdc As Long, ByVal nMapMode As Long) As Long
Private Declare Function SetWindowExtEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpSize As Size) As Long
Private Declare Function SetWindowOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nX As Long, ByVal nY As Long, lpPoint As POINTAPI) As Long
Private Declare Function SaveDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function RestoreDC Lib "gdi32" (ByVal hdc As Long, ByVal nSavedDC As Long) As Long
' These APIs are used to BitBlt the bitmap image into the metafile
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

' Used for creating the temporary WMF file
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Const MM_ANISOTROPIC = 8 ' Map mode anisotropic

Public Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hdc As Long, graphics As Long) As GpStatus
Public Declare Function GdipCreateFromHWND Lib "gdiplus" (ByVal hwnd As Long, graphics As Long) As GpStatus
Public Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As Long) As GpStatus
Public Declare Function GdipGetDC Lib "gdiplus" (ByVal graphics As Long, hdc As Long) As GpStatus
Public Declare Function GdipReleaseDC Lib "gdiplus" (ByVal graphics As Long, ByVal hdc As Long) As GpStatus
Public Declare Function GdipDrawImageRect Lib "gdiplus" (ByVal graphics As Long, ByVal image As Long, ByVal X As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single) As GpStatus
Public Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal filename As String, image As Long) As GpStatus
Public Declare Function GdipCloneImage Lib "gdiplus" (ByVal image As Long, cloneImage As Long) As GpStatus
Public Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal image As Long, Width As Long) As GpStatus
Public Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal image As Long, Height As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hbm As Long, ByVal hpal As Long, bitmap As Long) As GpStatus
Public Declare Function GdipBitmapGetPixel Lib "gdiplus" (ByVal bitmap As Long, ByVal X As Long, ByVal y As Long, color As Long) As GpStatus
Public Declare Function GdipBitmapSetPixel Lib "gdiplus" (ByVal bitmap As Long, ByVal X As Long, ByVal y As Long, ByVal color As Long) As GpStatus
Public Declare Function GdipDisposeImage Lib "gdiplus" (ByVal image As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromFile Lib "gdiplus" (ByVal filename As Long, bitmap As Long) As GpStatus

Public Type GdiplusStartupInput
   GdiplusVersion As Long              ' Must be 1 for GDI+ v1.0, the current version as of this writing.
   DebugEventCallback As Long          ' Ignored on free builds
   SuppressBackgroundThread As Long    ' FALSE unless you're prepared to call
                                       ' the hook/unhook functions properly
   SuppressExternalCodecs As Long      ' FALSE unless you want GDI+ only to use
                                       ' its internal image codecs.
End Type


Public Declare Function GdiplusStartup Lib "gdiplus" (token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As GpStatus
Public Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal token As Long)

Public Enum GpStatus   ' aka Status
   Ok = 0
   GenericError = 1
   InvalidParameter = 2
   OutOfMemory = 3
   ObjectBusy = 4
   InsufficientBuffer = 5
   NotImplemented = 6
   Win32Error = 7
   WrongState = 8
   Aborted = 9
   FileNotFound = 10
   ValueOverflow = 11
   AccessDenied = 12
   UnknownImageFormat = 13
   FontFamilyNotFound = 14
   FontStyleNotFound = 15
   NotTrueTypeFont = 16
   UnsupportedGdiplusVersion = 17
   GdiplusNotInitialized = 18
   PropertyNotFound = 19
   PropertyNotSupported = 20
End Enum


Public Function MakeRTBTransparent(RTBCtl As Object) As Boolean
' This function was taken from UDMX IOCP©'s
' Custom Rich Text box. It makes a Rich Text Box
' transparent, which is excellent if you slip a picture
' behind it.
On Error Resume Next
RTBCtl.BackColor = RTBCtl.Parent.BackColor
SetWindowLong RTBCtl.hwnd, GWL_EXSTYLE, WS_EX_TRANSPARENT
MakeRTBTransparent = Err.LastDllError = 0

End Function


Public Function StdPicAsRTF(aStdPic As StdPicture) As String
    Dim hMetaDC     As Long
    Dim hMeta       As Long
    Dim hPicDC      As Long
    Dim hOldBmp     As Long
    Dim aBMP        As bitmap
    Dim aSize       As Size
    Dim aPt         As POINTAPI
    Dim filename    As String
'    Dim aMetaHdr    As METAHEADER
    Dim screenDC    As Long
    Dim headerStr   As String
    Dim retStr      As String
    Dim byteStr     As String
    Dim bytes()     As Byte
    Dim filenum     As Integer
    Dim numBytes    As Long
    Dim i           As Long
    
    ' Create a metafile to a temporary file in the registered windows TEMP folder
    filename = getTempName("WMF")
    hMetaDC = CreateMetaFile(filename)
    
    ' Set the map mode to MM_ANISOTROPIC
    SetMapMode hMetaDC, MM_ANISOTROPIC
    ' Set the metafile origin as 0, 0
    SetWindowOrgEx hMetaDC, 0, 0, aPt
    ' Get the bitmap's dimensions
    GetObject aStdPic.handle, Len(aBMP), aBMP
    ' Set the metafile width and height
    SetWindowExtEx hMetaDC, aBMP.bmWidth, aBMP.bmHeight, aSize
    ' save the new dimensions
    SaveDC hMetaDC
    ' OK. Now transfer the freakin image to the metafile
    screenDC = GetDC(0)
    hPicDC = CreateCompatibleDC(screenDC)
    ReleaseDC 0, screenDC
    hOldBmp = SelectObject(hPicDC, aStdPic.handle)
    BitBlt hMetaDC, 0, 0, aBMP.bmWidth, aBMP.bmHeight, hPicDC, 0, 0, vbSrcCopy
    SelectObject hPicDC, hOldBmp
    DeleteDC hPicDC
    DeleteObject hOldBmp
    ' "redraw" the metafile DC
    RestoreDC hMetaDC, True
    ' close it and get the metafile handle
    hMeta = CloseMetaFile(hMetaDC)
    
'    GetObject hMeta, Len(aMetaHdr), aMetaHdr
    ' delete it from memory
    DeleteMetaFile hMeta
    
    ' Do the RTF header for the object. This little bit is sometimes required on
    '  earlier versions of the rich text box and in certain operating systems
    '  (WinNT springs to mind)
    headerStr = "{\rtf1\ansi"
    ' Picture specific tag stuff
    headerStr = headerStr & _
                "{\pict\picscalex100\picscaley100" & _
                "\picw" & aStdPic.Width & "\pich" & aStdPic.Height & _
                "\picwgoal" & aBMP.bmWidth * Screen.TwipsPerPixelX & _
                "\pichgoal" & aBMP.bmHeight * Screen.TwipsPerPixelY & _
                "\wmetafile8"
    
    ' Get the size of the metafile
    numBytes = FileLen(filename)
    ' Create our byte buffer for reading
    ReDim bytes(1 To numBytes)
    ' get a free file number
    filenum = FreeFile()
    ' open the file for input
    Open filename For Binary Access Read As #filenum
    ' read the bytes
    Get #filenum, , bytes
    ' close the file
    Close #filenum
    ' Generate our hex encoded byte string
    byteStr = String(numBytes * 2, "0")
    For i = LBound(bytes) To UBound(bytes)
        If bytes(i) > &HF Then
            Mid$(byteStr, 1 + (i - 1) * 2, 2) = Hex$(bytes(i))
        Else
            Mid$(byteStr, 2 + (i - 1) * 2, 1) = Hex$(bytes(i))
        End If
    Next i
    ' stick it all together
    retStr = headerStr & " " & byteStr & "}"
    ' Add in the closing RTF bit
    retStr = retStr & "}"
        
    StdPicAsRTF = retStr
    
    On Local Error Resume Next
    ' Kill the temporary file
    If Dir(filename) <> "" Then Kill filename
End Function

Public Sub AddPic(stdPic As StdPicture, Box As RichTextBox)
' A simple sub to add a WMF picture into a rich
' text box. Simple, and effective
    Box.SelRTF = modImages.StdPicAsRTF(stdPic)
End Sub

Private Function getTempName(Optional anExt As String = "tmp") As String
' This retrieves the temp path on your drive
' eg. c:\windows\temp
    Dim tempPath    As String
    Dim filename    As String
    Dim i           As Long
    
    Const validChars As String = "123567890qwertyuiopasdfghjklzxcvbnm"
    
    ' Create a buffer
    tempPath = String$(255, " ")
    ' get the system path
    GetTempPath 255, tempPath
    ' trim off the fat
    tempPath = Left$(tempPath, InStr(tempPath, Chr$(0)) - 1)
    ' Create a buffer
    filename = Space(12)
    ' Put the non-random stuff into the string
    Mid$(filename, 1, 1) = "T"
    Mid$(filename, Len(filename) - Len(anExt), 1) = "."
    ' Add in the specified extension, if provided ("tmp" is default)
    Mid$(filename, Len(filename) - Len(anExt) + 1, Len(anExt)) = anExt
    ' fill the buffer with random stuff
    Randomize
    For i = 2 To Len(filename) - 4
        Mid$(filename, i, 1) = Mid$(validChars, CLng(Rnd() * (Len(validChars)) + 1), 1)
    Next i
    tempPath = tempPath & filename
    ' return the path name
    getTempName = tempPath
    
End Function

Public Function SendPic(Pic As StdPicture)
' Allows us to send a picture via winsock, by converting
' a picture into RTF code, then sending it over sckUDP.
' It's slow, it's horrible, and it's buggy, but it works!
Dim X As Boolean
Dim tempint As Long
Dim tempintB As Long
Dim T As String
T = StdPicAsRTF(Pic)
tempintB = 1
tempint = 1
Broadcast "startpic"
DoEvents
DoEvents

Do Until X = True
If tempint > Len(T) Then X = True
frmLog.Label1.Caption = "Picture Sent: " & tempint & " / " & Len(T)
Broadcast "thepicø" & Trim(Mid(T, tempint, 200))
tempintB = temintb + 1
tempint = tempint + 200
DoEvents
DoEvents
Loop
DoEvents
DoEvents
Broadcast "endpic"
End Function
