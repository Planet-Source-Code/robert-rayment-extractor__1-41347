Attribute VB_Name = "ApiDibs"
'ApiDibs.bas


Option Base 1  ' Arrays base 1
DefLng A-W     ' All variable Long
DefSng X-Z     ' unless singles
               ' unless otherwise defined
        
' -----------------------------------------------------------
' API to resize to a rectangle
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, _
ByVal X As Long, ByVal Y As Long, _
ByVal nWidth As Long, ByVal nHeight As Long, _
ByVal hSrcDC As Long, _
ByVal xSrc As Long, ByVal ySrc As Long, _
ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, _
ByVal dwRop As Long) As Long
        
' -----------------------------------------------------------
' API to Fill background

Public Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, _
 ByVal Y As Long, ByVal crColor As Long, ByVal fuFillType As Long) As Long

' --------------------------------------------------------------
' Windows API - For blitting one image to another location

Public Declare Function BitBlt Lib "gdi32" _
(ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, _
ByVal nWidth As Long, ByVal nHeight As Long, _
ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, _
ByVal dwRop As Long) As Long

' --------------------------------------------------------------
' Function & constants to make Window stay on top

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
ByVal hwndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, _
ByVal wi As Long, ByVal ht As Long, ByVal wflags As Long) As Long

Public Const hwndInsertAfter = -1
Public Const wflags = &H40 Or &H20

'--------------------------------------------------------------------------
' Shaping APIs

Public Declare Function CreateRoundRectRgn Lib "gdi32" _
(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long

Public Declare Function SetWindowRgn Lib "user32" _
(ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Public Declare Function DeleteObject Lib "gdi32" _
(ByVal hObject As Long) As Long

               
' -----------------------------------------------------------
'  This required instead of Screen.Height & Width for resizing

Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Const SM_CXSCREEN = 0 'X Size of screen
Public Const SM_CYSCREEN = 1 'Y Size of Screen

' -----------------------------------------------------------
Public Declare Function GetPixel Lib "gdi32" _
(ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

' -----------------------------------------------------------
' APIs for getting DIB bits to PicMem

Public Declare Function GetDIBits Lib "gdi32" _
(ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, _
ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long

Public Declare Function CreateCompatibleDC Lib "gdi32" _
(ByVal hdc As Long) As Long

Public Declare Function SelectObject Lib "gdi32" _
(ByVal hdc As Long, ByVal hObject As Long) As Long

Public Declare Function DeleteDC Lib "gdi32" _
(ByVal hdc As Long) As Long

'------------------------------------------------------------------------------

'To fill BITMAP structure
Public Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" _
(ByVal hObject As Long, ByVal Lenbmp As Long, dimbmp As Any) As Long

Public Type BITMAP
   bmType As Long              ' Type of bitmap
   bmWidth As Long             ' Pixel width
   bmHeight As Long            ' Pixel height
   bmWidthBytes As Long        ' Byte width = 4 x Pixel width here
   bmPlanes As Integer         ' Color depth of bitmap
   bmBitsPixel As Integer      ' Bits per pixel, must be 16 or 24
   bmBits As Long              ' This is the pointer to the bitmap data  !!!
End Type
Public bmp As BITMAP

'------------------------------------------------------------------------------

' Structures for StretchDIBits
Public Type BITMAPINFOHEADER ' 40 bytes
   biSize As Long
   biwidth As Long
   biheight As Long
   biPlanes As Integer
   biBitCount As Integer
   biCompression As Long
   biSizeImage As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed As Long
   biClrImportant As Long
End Type

Public Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type

Public Type BITMAPINFO
   bmiH As BITMAPINFOHEADER
   Colors(0 To 1) As RGBQUAD
End Type
Public bm As BITMAPINFO

' For transferring drawing in an array to Form or PicBox 7 printing
Public Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, _
ByVal X As Long, ByVal Y As Long, _
ByVal DesW As Long, ByVal DesH As Long, _
ByVal SrcXOffset As Long, ByVal SrcYOffset As Long, _
ByVal PICWW As Long, ByVal PICHH As Long, _
lpBits As Any, lpBitsInfo As BITMAPINFO, _
ByVal wUsage As Long, ByVal dwRop As Long) As Long

'wUsage is one of:-
Public Const DIB_PAL_COLORS = 1 '  uses system
Public Const DIB_RGB_COLORS = 0 '  uses RGBQUAD
'dwRop is vbSrcCopy
'------------------------------------------------------------------------------

' To invert picture box
Public Declare Function SetRect Lib "user32" (lpRect As RECT, _
ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Public Declare Function InvertRect Lib "user32" _
(ByVal hdc As Long, lpRect As RECT) As Long

Public Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Public IR As RECT
'------------------------------------------------------------------------------


'Public pic1mem() As Byte  ' Holds loaded picture
'Public pic2mem() As Byte  ' Holds dithered picture

Public Sub FillBMPStruc(ByVal bwidth, ByVal bheight)
  
  With bm.bmiH
   .biSize = 40
   .biwidth = bwidth
   .biheight = bheight
   .biPlanes = 1
   .biBitCount = 32     ' always 32 in this prog
   .biCompression = 0
   BScanLine = (((bwidth * .biBitCount) + 31) \ 32) * 4
   ' Ensure expansion to 4B boundary
   BScanLine = (Int((BScanLine + 3) \ 4)) * 4

   .biSizeImage = BScanLine * Abs(.biheight)
   .biXPelsPerMeter = 0
   .biYPelsPerMeter = 0
   .biClrUsed = 0
   .biClrImportant = 0
 End With

End Sub

Public Sub GETDIBS(ByVal PICIM As Long)

' PICIM is picbox.Image - handle to picbox memory
' from which pixels will be extracted and
' stored in pic1mem()

On Error GoTo DIBError

'Get info on picture loaded into PIC
GetObjectAPI PICIM, Len(bmp), bmp

NewDC = CreateCompatibleDC(0&)
OldH = SelectObject(NewDC, PICIM)

FillBMPStruc bmp.bmWidth, bmp.bmHeight

' Set PicMem to receive color bytes or indexes or bits
picHt = bmp.bmHeight
picWd = bmp.bmWidth
ReDim pic1mem(4, picWd, picHt)
ReDim pic2mem(4, picWd, picHt)

' Load color bytes to PicMem
ret = GetDIBits(NewDC, PICIM, 0, picHt, pic1mem(1, 1, 1), bm, 1)

' Clear mem
SelectObject NewDC, OldH
DeleteDC NewDC

DStruc.Ptrpic1mem = VarPtr(pic1mem(1, 1, 1))
DStruc.Ptrpic2mem = VarPtr(pic2mem(1, 1, 1))

Exit Sub
'==========
DIBError:
  MsgBox "DIB Error in GETDIBS", , "Extractor"
  DoEvents
  Unload Form1
  End
End Sub

