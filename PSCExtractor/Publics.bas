Attribute VB_Name = "Publics"
' Publics.bas

Option Base 1

DefLng A-W
DefSng X-Z

Public red As Byte, green As Byte, blue As Byte
' Same & RGB +/- limits
Public DELRGB()

Public pic1mem() As Byte   ' For Loaded picture
Public pic2mem() As Byte   ' For Output dither picture
Public picHt, picWd        ' Picture Height & Width

Public BScanLine     ' For saving file

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Ptrpic1mem '= VarPtr(pic1mem(1, 1, 1))
Public Ptrpic2mem '= VarPtr(pic2mem(1, 1, 1))
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public StartColor
Public PtrDELRGB

' VB<->ASMVB switch
Public ASMVB As Boolean

' DRAW MODE VARIABLES
Public TestDrawMode As Boolean   ' True Test/False Draw
Public DrawingMode As Boolean
Public DrawOpt                   ' 0 Extract rect, 1 Extract shape
                                 ' 2 Resize to rect
Public iXp, iYp, iX2, iY2
Public RectWidth, RectHeight


Public Sub LngToRGB(LCul)
'Convert Long Colors() to RGB components
red = (LCul And &HFF&)
green = (LCul And &HFF00&) / &H100&
blue = (LCul And &HFF0000) / &H10000
End Sub

Public Sub TRANSFER(frm As Form)
frm.Frame2.Visible = False

Screen.MousePointer = vbHourglass
DoEvents

frm.cmdPrint(0).Enabled = True
frm.cmdSavePic(0).Enabled = True
frm.cmdSavePic(1).Enabled = True

'InitBits
bm.Colors(0).rgbBlue = 0
bm.Colors(0).rgbGreen = 0
bm.Colors(0).rgbRed = 0
bm.Colors(1).rgbBlue = 255
bm.Colors(1).rgbGreen = 255
bm.Colors(1).rgbRed = 255

ReDim pic2mem(4, picWd, picHt)

DoEvents

If ASMVB Then
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
   DStruc.picHt = picHt
   DStruc.picWd = picWd
'   U1 = UBound(pic1mem, 1)
'   U2 = UBound(pic1mem, 2)
'   U3 = UBound(pic1mem, 3)
   
   DStruc.Ptrpic1mem = VarPtr(pic1mem(1, 1, 1))
   DStruc.Ptrpic2mem = VarPtr(pic2mem(1, 1, 1))
   
   DStruc.StartColor = StartColor
   DStruc.PtrDELRGB = VarPtr(DELRGB(0))
   ASM_Extract
   ShowDitheredPicture Form1
   Screen.MousePointer = vbDefault
   Exit Sub
   '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
End If

' VB VB VB
' NB NB Code written to match mcode approx.

LngToRGB StartColor
R1 = red
G1 = green
B1 = blue

R1minus = R1 - DELRGB(1)
R1plus = R1 + DELRGB(1)
G1minus = G1 - DELRGB(2)
G1plus = G1 + DELRGB(2)
B1minus = B1 - DELRGB(3)
B1plus = B1 + DELRGB(3)

For j = 1 To picHt
For i = 1 To picWd
   
   B = pic1mem(1, i, j)
   G = pic1mem(2, i, j)
   R = pic1mem(3, i, j)
   
   If R < R1minus Then GoTo Whiten
   If R > R1plus Then GoTo Whiten
   If G < G1minus Then GoTo Whiten
   If G > G1plus Then GoTo Whiten
   If B < B1minus Then GoTo Whiten
   If B > B1plus Then GoTo Whiten
   red = 255 '0
   GoTo Fillpic2
Whiten:
   red = 0 '255
Fillpic2:
   pic2mem(1, i, j) = red
   pic2mem(2, i, j) = red
   pic2mem(3, i, j) = red

Nexi:
Next i
Next j

ShowDitheredPicture Form1
Screen.MousePointer = vbDefault
End Sub

Public Sub ShowDitheredPicture(frm As Form)

'Public Const DIB_PAL_COLORS = 1 '  uses system colors
'Public Const DIB_RGB_COLORS = 0 '  uses RGBQUAD colors

If StretchDIBits(frm.pic2.hdc, _
   0, 0, picWd, picHt, _
   0, 0, picWd, picHt, _
   pic2mem(1, 1, 1), bm, _
   DIB_RGB_COLORS, vbSrcCopy) = 0 Then
      MsgBox "Blit Error", , "Extractor"
      Done = True
      Erase pic1mem, pic2mem
      Unload frm
      End
End If
frm.pic2.Refresh
End Sub


