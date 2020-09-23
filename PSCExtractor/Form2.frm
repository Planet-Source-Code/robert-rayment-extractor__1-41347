VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00008000&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Test mask"
   ClientHeight    =   6090
   ClientLeft      =   3060
   ClientTop       =   1200
   ClientWidth     =   4965
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   406
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   331
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraTestDraw 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   75
      TabIndex        =   15
      Top             =   60
      Width           =   3465
      Begin VB.OptionButton optDraw 
         BackColor       =   &H00FFFF00&
         Height          =   315
         Index           =   2
         Left            =   1620
         Picture         =   "Form2.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Resize to rectangle"
         Top             =   30
         Width           =   315
      End
      Begin VB.OptionButton optDraw 
         BackColor       =   &H00FFFF00&
         Height          =   315
         Index           =   1
         Left            =   1245
         Picture         =   "Form2.frx":00D2
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Extract shape"
         Top             =   30
         Width           =   315
      End
      Begin VB.OptionButton optDraw 
         BackColor       =   &H00FFFF00&
         Height          =   315
         Index           =   0
         Left            =   870
         Picture         =   "Form2.frx":01A4
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Extract rectangle"
         Top             =   30
         Width           =   315
      End
      Begin VB.CommandButton cmdDraw 
         BackColor       =   &H00C0C000&
         Height          =   315
         Index           =   1
         Left            =   405
         Picture         =   "Form2.frx":0276
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Undo"
         Top             =   30
         Width           =   315
      End
      Begin VB.CommandButton cmdDraw 
         BackColor       =   &H00C0C000&
         Height          =   315
         Index           =   0
         Left            =   45
         Picture         =   "Form2.frx":0348
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Proceed"
         Top             =   30
         Width           =   315
      End
      Begin VB.Label LabWH 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         Caption         =   "     W      H"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2925
         TabIndex        =   22
         Top             =   15
         Width           =   480
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         Caption         =   "    M-Dn  Draw     M-Up   Done"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1980
         TabIndex        =   18
         Top             =   30
         Width           =   990
      End
   End
   Begin VB.Timer Timer1 
      Left            =   30
      Top             =   5625
   End
   Begin VB.CommandButton cmdShiftFocus 
      Caption         =   "Command1"
      Height          =   195
      Left            =   165
      TabIndex        =   13
      Top             =   5565
      Width           =   150
   End
   Begin VB.Frame fraScr 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   390
      Index           =   1
      Left            =   645
      TabIndex        =   4
      Top             =   5565
      Width           =   3945
      Begin VB.CommandButton cmdScr 
         BackColor       =   &H0000C000&
         Height          =   225
         Index           =   3
         Left            =   3570
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   90
         Width           =   225
      End
      Begin VB.CommandButton cmdScr 
         BackColor       =   &H0000C000&
         Height          =   225
         Index           =   2
         Left            =   225
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   75
         Width           =   225
      End
      Begin VB.Label LabThumb 
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Index           =   1
         Left            =   1800
         TabIndex        =   12
         Top             =   75
         Width           =   180
      End
      Begin VB.Label LabBack 
         BackColor       =   &H0000C000&
         Height          =   270
         Index           =   1
         Left            =   660
         TabIndex        =   10
         Top             =   90
         Width           =   2655
      End
   End
   Begin VB.Frame fraScr 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2595
      Index           =   0
      Left            =   30
      TabIndex        =   2
      Top             =   375
      Width           =   330
      Begin VB.CommandButton cmdScr 
         BackColor       =   &H0000C000&
         Height          =   225
         Index           =   1
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1830
         Width           =   225
      End
      Begin VB.CommandButton cmdScr 
         BackColor       =   &H0000C000&
         Height          =   225
         Index           =   0
         Left            =   45
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   45
         Width           =   225
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   15
         Left            =   120
         TabIndex        =   3
         Top             =   1950
         Width           =   30
      End
      Begin VB.Label LabThumb 
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Index           =   0
         Left            =   60
         TabIndex        =   11
         Top             =   855
         Width           =   180
      End
      Begin VB.Label LabBack 
         BackColor       =   &H0000C000&
         Height          =   1410
         Index           =   0
         Left            =   45
         TabIndex        =   9
         Top             =   330
         Width           =   225
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00008000&
      Height          =   5130
      Left            =   450
      ScaleHeight     =   338
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   253
      TabIndex        =   0
      Top             =   450
      Width           =   3855
      Begin VB.PictureBox picBack 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0080FFFF&
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   750
         ScaleHeight     =   97
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   116
         TabIndex        =   14
         Top             =   570
         Visible         =   0   'False
         Width           =   1740
      End
      Begin VB.PictureBox picShow 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   0
         ScaleHeight     =   97
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   116
         TabIndex        =   1
         Top             =   -15
         Width           =   1740
         Begin VB.Shape Shape1 
            BorderColor     =   &H00C0FFFF&
            FillStyle       =   6  'Cross
            Height          =   180
            Left            =   135
            Top             =   135
            Width           =   180
         End
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Form2

' Extractors

Option Base 1

DefLng A-W  ' Longs
DefSng X-Z  ' Singles

Private Sub cmdDraw_Click(Index As Integer)
If Index = 0 Then
   'PROCEED
   If DrawOpt = 0 Then
      'EXTRACT RECTANGLE
      Shape1.Visible = False
      With Form1.pic1
         .Width = RectWidth
         .Height = RectHeight
      End With
      With Form1.pic2
         .Width = RectWidth
         .Height = RectHeight
      End With
      Form1.pic1.Picture = LoadPicture
      Form1.pic2.Picture = LoadPicture
      
      res = StretchBlt(Form1.pic1.hdc, 0, 0, RectWidth, RectHeight, _
      picShow.hdc, iXp, iYp, RectWidth, RectHeight, vbSrcCopy)
      
      picWd = RectWidth
      picHt = RectHeight
      
      Form1.pic1.Refresh
      Form1.pic2.Refresh
   
   ElseIf DrawOpt = 2 Then
      'RESIZE TO RECTANGLE
      Shape1.Visible = False
      With Form1.pic1
         .Width = RectWidth
         .Height = RectHeight
      End With
      With Form1.pic2
         .Width = RectWidth
         .Height = RectHeight
      End With
      Form1.pic1.Picture = LoadPicture
      Form1.pic2.Picture = LoadPicture
      
      res = StretchBlt(Form1.pic1.hdc, 0, 0, RectWidth, RectHeight, _
      picShow.hdc, 0, 0, picWd, picHt, vbSrcCopy)

      picWd = RectWidth
      picHt = RectHeight
      
      Form1.pic1.Refresh
      Form1.pic2.Refresh
   
   Else
      'EXTRACT SHAPE
      Shape1.Visible = False
      picBack.DrawStyle = vbSolid
      picBack.DrawMode = 13
      picBack.DrawWidth = 1
      picBack.FillColor = vbWhite
      picBack.FillStyle = vbFSSolid
      pX = 1: pY = 1
      If picBack.Point(pX, pY) <> vbBlack Then
         For i = 1 To 10
            pX = pX + 1: pY = pY + 1
            If picBack.Point(pX, pY) = vbBlack Then Exit For
         Next
         If i = 11 Then ' Black not found
            picBack.FillStyle = vbTransparent  'Default (Transparent)
            Unload Form2
            MsgBox "Can't make mask.  Try drawing again", , "Extractor"
            Exit Sub
         End If
      End If
      FillPtcul& = vbBlack 'vbWhite 'picBack.Point(pX,pY)
      FLOODFILLSURFACE = 1
      'Fills with FillColor so long as point surrounded by FillPtcul&
      rs = ExtFloodFill(picBack.hdc, pX, pY, FillPtcul&, FLOODFILLSURFACE)
      picBack.Refresh
      picBack.FillStyle = vbTransparent  'Default (Transparent)
   
      ' Invert mask
      res = SetRect(IR, 0, 0, picWd, picHt)
      res = InvertRect(picBack.hdc, IR)
      picBack.Refresh
      ' Copy to pic2
      BitBlt Form1.pic2.hdc, 0, 0, picWd, picHt, picBack.hdc, 0, 0, vbSrcCopy
      Form1.pic2.Refresh
      TestMask
      ' Copy to pic1
      BitBlt Form1.pic1.hdc, 0, 0, picWd, picHt, picShow.hdc, 0, 0, vbSrcCopy
      Form1.pic1.Refresh
      
   End If
   
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      GETDIBS Form1.pic1.Image  ' BPP Fills pic1mem and sizes pic2mem
      Form1.cmdTest.Enabled = True
      Form1.cmdPrint(0).Enabled = True
      Form1.cmdSavePic(0).Enabled = True
      Form1.cmdSavePic(1).Enabled = True
      Form1.Frame1.Visible = False
      
      Form_Unload 0

Else
   'UNDO
   Shape1.Visible = False
   LabWH.Caption = ""
   cmdDraw(0).Enabled = False
   BlitPicture
End If

End Sub

Private Sub Form_Load()
''
optDraw(0).Value = True
DrawOpt = 0    ' Rectangle
Shape1.Visible = False
DrawingMode = False
cmdDraw(0).Enabled = False
End Sub

Private Sub Form_Resize()

'----------------------------------------------
If Form2.Width < 3990 Then Form2.Width = 3990
If Form2.Height < 3990 Then Form2.Height = 3990

' Make form stay on top X,Y,WI,HI
FT = Form1.Top / Screen.TwipsPerPixelY + 60
ret& = SetWindowPos(Me.hwnd, hwndInsertAfter, _
300, FT, Form2.Width \ 15, Form2.Height \ 15, wflags)


p1w = Form2.Width \ 15 - 60
p1h = Form2.Height \ 15 - 60 - 30
With Picture1
   .Left = 30
   .Top = 30
   .Width = p1w
   .Height = p1h
End With
With picShow
   .Left = 0
   .Top = 0
   .Width = picWd
   .Height = picHt
End With

With picBack
   '.Left = 0
   '.Top = 0
   .Width = picWd
   .Height = picHt
End With

'--- Default scroll bar values ----------------
'--- Vertical Scroll Bar Settings -------------
scrMax(0) = picHt:
scrMin(0) = 0
SmallChange(0) = 1
LargeChange(0) = 1 ' NB > 0
'--- Horizontal Scroll Bar Settings -----------
scrMax(1) = picWd:
scrMin(1) = 0
SmallChange(1) = 1
LargeChange(1) = 1 ' NB > 0
'----------------------------------------------
' SET UP SCROLL BARS
With fraScr(0)
   .Left = 10
   .Top = Picture1.Top + 2
   .Width = 20
   .Height = Picture1.Height
End With

Picture1.Left = fraScr(0).Left + fraScr(0).Width + 5
picShow.Left = 0

With fraScr(1)
   .Top = Picture1.Top + Picture1.Height + 2
   .Left = Picture1.Left
   .Height = 20
   .Width = Picture1.Width
End With

cmdShiftFocus.Visible = False ' Shift focus from cmdScr(0,1/2,3)

If picHt > Picture1.Height Then
   fraScr(0).Visible = True
   scrMax(0) = picHt - Form2.Picture1.Height + 6
   LargeChange(0) = scrMax(0) \ 10
   If LargeChange(0) < 1 Then LargeChange(0) = 1
Else
   fraScr(0).Visible = False
End If

If picWd > Picture1.Width Then
   fraScr(1).Visible = True
   scrMax(1) = picWd - Form2.Picture1.Width + 6
   LargeChange(1) = scrMax(1) \ 10
   If LargeChange(1) < 1 Then LargeChange(1) = 1
Else
   fraScr(1).Visible = False
End If

ModReDim Form2
nTVal(0) = scrMin(0)
nTVal(1) = scrMin(1)

' Initial scroll bar values
VVal = nTVal(0)
HVal = nTVal(1)
'----------------------------------------------

' TEST MASK
If TestDrawMode = True Then
   fraTestDraw.Visible = False
   Caption = "TEST MASK"
   
   TestMask
'----------------------------------------------
' DRAW MODE
Else
   fraTestDraw.Visible = True
   Caption = "EXTRACTOR"
   
   BlitPicture
End If

'----------------------------------------------
End Sub

Private Sub TestMask()
   
   picShow.Picture = LoadPicture
   picShow.Refresh
   picBack.Picture = LoadPicture
   picBack.Refresh
   ' Make PicBack = picShow
   picShow.BackColor = 0
   picBack.BackColor = picShow.BackColor
   BitBlt picBack.hdc, 0, 0, picWd, picHt, picShow, 0, 0, vbSrcCopy
   ' NB picBack has same background as picShow
   
   ' Invert mask
   res = SetRect(IR, 0, 0, picWd, picHt)
   res = InvertRect(Form1.pic2.hdc, IR)
   Form1.pic2.Refresh
   ' And inverted mask with picback background
   BitBlt picBack.hdc, 0, 0, picWd, picHt, Form1.pic2.hdc, 0, 0, vbSrcAnd
   ' Reset mask
   res = SetRect(IR, 0, 0, picWd, picHt)
   res = InvertRect(Form1.pic2.hdc, IR)
   Form1.pic2.Refresh
   
   ' Copy org picture
   BitBlt picShow.hdc, 0, 0, picWd, picHt, Form1.pic1.hdc, 0, 0, vbSrcCopy
   ' And in mask
   BitBlt picShow.hdc, 0, 0, picWd, picHt, Form1.pic2.hdc, 0, 0, vbSrcAnd
   picShow.Picture = picShow.Image
   
   ' Or in pickback to show green background instead of black
   BitBlt picShow.hdc, 0, 0, picWd, picHt, picBack.hdc, 0, 0, vbSrcPaint

End Sub

Private Sub BlitPicture()
   picShow.Picture = LoadPicture
   picShow.Refresh
   picBack.Picture = LoadPicture
   picBack.Refresh
   ' Make PicBack = picShow
   picShow.BackColor = 0
   picBack.BackColor = picShow.BackColor
   BitBlt picShow.hdc, 0, 0, picWd, picHt, Form1.pic1.hdc, 0, 0, vbSrcCopy
   picShow.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Form2
DoEvents

' Reposition main pictures
ModReDim Form1
nTVal(0) = scrMin(0)
nTVal(1) = scrMin(1)
' Vert
Form1.pic1.Top = -nTVal(0)
Form1.pic1.Refresh
Form1.pic2.Top = -nTVal(0)
Form1.pic2.Refresh
' Horz
Form1.pic1.Left = -nTVal(1)
Form1.pic1.Refresh
Form1.pic2.Left = -nTVal(1)
Form1.pic2.Refresh

End Sub

'#### DRAWING ##############################################

Private Sub optDraw_Click(Index As Integer)
DrawOpt = Index
' 0 Extract rect
' 1 Extract shape
' 2 Resize to rect
End Sub


Private Sub picShow_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

BlitPicture
cmdDraw(0).Enabled = True
Shape1.Visible = False
DoEvents
LabWH.Caption = ""

DrawingMode = True

If DrawOpt = 0 Or DrawOpt = 2 Then
   'RECTANGLE OR RESIZE
   iXp = X: iYp = Y
   iX2 = X: iY2 = Y
ElseIf DrawOpt = 1 Then
   'SHAPE
   picShow.DrawMode = vbXorPen
   iXp = X: iYp = Y
   iX2 = X: iY2 = Y
   picShow.PSet (iXp, iYp), vbWhite
   picBack.PSet (iXp, iYp), vbWhite
Else
   DrawingMode = False
   Exit Sub
End If

End Sub

Private Sub picShow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If DrawingMode Then

   If DrawOpt = 0 Or DrawOpt = 2 Then
      'RECTANGLE OR RESIZE
      iX2 = X
      iY2 = Y
      Shape1.Visible = True
      Shape1.Top = iYp
      Shape1.Left = iXp
      If iX2 <= iXp Then iX2 = iXp
      Shape1.Width = iX2 - iXp
      If iY2 <= iYp Then iY2 = iYp
      Shape1.Height = iY2 - iYp
      LabWH.Caption = Str$(iX2 - iXp) & vbLf & Str$(iY2 - iYp)
   Else
      'SHAPE
      If Shift = 1 Then    ' Shift key X fixed
         iY2 = Y
      ElseIf Shift = 2 Then   ' Ctrl key Y fixed
         iX2 = X
      Else
         iX2 = X
         iY2 = Y
      End If
      
      picShow.Line -(iX2, iY2), vbWhite
      picBack.Line -(iX2, iY2), vbWhite
   End If

End If

End Sub

Private Sub picShow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If DrawingMode Then

   If DrawOpt = 0 Or DrawOpt = 2 Then
      'RECTANGLE OR RESIZE
      iX2 = X
      iY2 = Y
      If iX2 < iXp Then iX2 = iXp
      RectWidth = iX2 - iXp
      If RectWidth < 16 Then
         RectWidth = 16
         iX2 = iXp + 16
      End If
      If iY2 < iYp Then iY2 = iYp
      RectHeight = iY2 - iYp
      If RectHeight < 16 Then
         RectHeight = 16
         iY2 = iYp + 16
      End If
      DrawingMode = False
   Else
      'SHAPE
      picShow.DrawWidth = 2
      picBack.DrawWidth = 2
      picShow.Line (iXp, iYp)-(iX2, iY2), vbWhite
      picBack.Line (iXp, iYp)-(iX2, iY2), vbWhite
      picShow.DrawMode = vbCopyPen
      DrawingMode = False
      picShow.DrawWidth = 1
      picBack.DrawWidth = 1
   End If

End If
End Sub


'###################################################################
'####### SCROLL BARS ###############################################
'###################################################################

Private Sub LabThumb_MouseDown(Index As Integer, Button As Integer, Shift As Integer, _
X As Single, Y As Single)
   LabThumbMouseDown Form2, Index, X, Y
   
   If nTVal(0) <> VVal Or nTVal(1) <> HVal Then
      MovePics
      VVal = nTVal(0)
      HVal = nTVal(1)
   End If
   
End Sub

Private Sub LabThumb_MouseMove(Index As Integer, Button As Integer, Shift As Integer, _
X As Single, Y As Single)
   LabThumbMouseMove Form2, Index, X, Y
   
   If nTVal(0) <> VVal Or nTVal(1) <> HVal Then
      MovePics
      VVal = nTVal(0)
      HVal = nTVal(1)
   End If

End Sub

Private Sub LabThumb_MouseUp(Index As Integer, Button As Integer, Shift As Integer, _
X As Single, Y As Single)
   ' Snap thumb position
   LabThumbMouseUp Form2, Index, X, Y
   
   If nTVal(0) <> VVal Or nTVal(1) <> HVal Then
      MovePics
      VVal = nTVal(0)
      HVal = nTVal(1)
   End If

End Sub

'###################################################################

Private Sub LabBack_MouseDown(Index As Integer, Button As Integer, Shift As Integer, _
X As Single, Y As Single)
   
   LabBackMouseDown Form2, Index, X, Y
   
   If nTVal(0) <> VVal Or nTVal(1) <> HVal Then
      MovePics
      VVal = nTVal(0)
      HVal = nTVal(1)
   End If

End Sub

Private Sub LabBack_MouseUp(Index As Integer, Button As Integer, Shift As Integer, _
X As Single, Y As Single)
   LabBackMouseUp Form2
   
   If nTVal(0) <> VVal Or nTVal(1) <> HVal Then
      MovePics
      VVal = nTVal(0)
      HVal = nTVal(1)
   End If

End Sub

'###################################################################

Private Sub cmdScr_MouseDown(Index As Integer, Button As Integer, Shift As Integer, _
X As Single, Y As Single)
   
   cmdScrMouseDown Form2, Index, Button, X, Y
End Sub

Private Sub cmdScr_MouseUp(Index As Integer, Button As Integer, Shift As Integer, _
X As Single, Y As Single)
   
   cmdScrMouseUp Form2
End Sub

Private Sub cmdScr_Click(Index As Integer)

   cmdScrMouseUp Form2

End Sub

'###################################################################

Private Sub Timer1_Timer()
   Timer1Timer Form2
   'LabVal(Index).Caption = nTVal(Index)  ' Out
   
   If nTVal(0) <> VVal Or nTVal(1) <> HVal Then
      MovePics
      VVal = nTVal(0)
      HVal = nTVal(1)
   End If

End Sub

'###################################################################
'###################################################################

Private Sub MovePics()
' Vert
picShow.Top = -nTVal(0)
picShow.Refresh
'' Horz
picShow.Left = -nTVal(1)
picShow.Refresh
End Sub



