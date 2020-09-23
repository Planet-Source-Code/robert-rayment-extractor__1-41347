VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00008000&
   BorderStyle     =   0  'None
   ClientHeight    =   7455
   ClientLeft      =   60
   ClientTop       =   -510
   ClientWidth     =   11730
   DrawWidth       =   2
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   497
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   782
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00008000&
      Height          =   2175
      Left            =   450
      TabIndex        =   16
      Top             =   345
      Width           =   1155
      Begin VB.CommandButton cmdSavePic 
         BackColor       =   &H80000009&
         Caption         =   "Save Pic"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   810
         Width           =   1035
      End
      Begin VB.CommandButton cmdExit 
         Height          =   345
         Index           =   0
         Left            =   405
         Picture         =   "Main.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "EXIT"
         Top             =   1650
         Width           =   390
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H80000009&
         Caption         =   "Print picture"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1125
         Width           =   1035
      End
      Begin VB.CommandButton cmdSavePic 
         BackColor       =   &H80000009&
         Caption         =   "Save Mask"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   495
         Width           =   1035
      End
      Begin VB.CommandButton cmdLoadPic 
         BackColor       =   &H80000009&
         Caption         =   "Load picture"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   180
         Width           =   1050
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00008000&
      Height          =   1260
      Left            =   1020
      TabIndex        =   54
      Top             =   660
      Width           =   1650
      Begin VB.CommandButton cmdDrawMode 
         BackColor       =   &H80000009&
         Caption         =   "Extractor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   870
         Width           =   1395
      End
      Begin VB.CommandButton cmdTest 
         BackColor       =   &H80000009&
         Caption         =   "Test Mask"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   555
         Width           =   1395
      End
      Begin VB.CommandButton cmdInvert 
         BackColor       =   &H80000009&
         Caption         =   "Invert Mask"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   240
         Width           =   1395
      End
   End
   Begin VB.CommandButton cmdTools 
      BackColor       =   &H80000009&
      Caption         =   "Tools"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1035
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   420
      Width           =   585
   End
   Begin VB.CommandButton cmdSetPM 
      BackColor       =   &H00E0E0E0&
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   5
      Left            =   7815
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   420
      Width           =   510
   End
   Begin VB.CommandButton cmdSetPM 
      BackColor       =   &H00E0E0E0&
      Caption         =   "50"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   4
      Left            =   7380
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   420
      Width           =   435
   End
   Begin VB.CommandButton cmdSetPM 
      BackColor       =   &H00E0E0E0&
      Caption         =   "30"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   3
      Left            =   6945
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   420
      Width           =   435
   End
   Begin VB.CommandButton cmdSetPM 
      BackColor       =   &H00E0E0E0&
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   2
      Left            =   6525
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   420
      Width           =   420
   End
   Begin VB.CommandButton cmdSetPM 
      BackColor       =   &H00E0E0E0&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   6105
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   420
      Width           =   420
   End
   Begin VB.CommandButton cmdSetPM 
      BackColor       =   &H00E0E0E0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   5775
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   420
      Width           =   330
   End
   Begin VB.PictureBox picLims 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   5010
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   70
      TabIndex        =   37
      Top             =   90
      Width           =   1050
   End
   Begin VB.PictureBox picLims 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   3435
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   70
      TabIndex        =   36
      Top             =   90
      Width           =   1050
   End
   Begin VB.PictureBox picLims 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   1815
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   70
      TabIndex        =   35
      Top             =   90
      Width           =   1050
   End
   Begin VB.CommandButton cmdVBASM 
      BackColor       =   &H00FFFFFF&
      Caption         =   "VB"
      Height          =   330
      Left            =   1035
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Toggle VB <-> ASM"
      Top             =   75
      Width           =   585
   End
   Begin VB.PictureBox picLims 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   6615
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   80
      TabIndex        =   24
      Top             =   90
      Width           =   1200
   End
   Begin VB.CommandButton cmdExit 
      Height          =   345
      Index           =   2
      Left            =   75
      Picture         =   "Main.frx":09C6
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "EXIT"
      Top             =   75
      Width           =   390
   End
   Begin VB.CommandButton cmdFile 
      BackColor       =   &H00FFFFFF&
      Caption         =   "File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   75
      Width           =   540
   End
   Begin VB.CommandButton cmdExit 
      Height          =   345
      Index           =   1
      Left            =   10935
      Picture         =   "Main.frx":0F4A
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "EXIT"
      Top             =   75
      Width           =   390
   End
   Begin VB.CheckBox chkResizer 
      BackColor       =   &H000000C0&
      Caption         =   "Resize"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "After screen res change"
      Top             =   420
      Width           =   915
   End
   Begin VB.Timer Timer1 
      Left            =   6285
      Top             =   6630
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00008000&
      Height          =   5625
      Left            =   6315
      ScaleHeight     =   371
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   345
      TabIndex        =   8
      Top             =   690
      Width           =   5235
      Begin VB.PictureBox pic2 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   705
         Left            =   120
         ScaleHeight     =   47
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   9
         Top             =   60
         Width           =   960
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00008000&
      Height          =   5685
      Left            =   495
      ScaleHeight     =   375
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   336
      TabIndex        =   6
      Top             =   735
      Width           =   5100
      Begin VB.Frame Frame3 
         BackColor       =   &H00008000&
         Height          =   1410
         Left            =   1050
         TabIndex        =   26
         Top             =   690
         Width           =   705
         Begin VB.CommandButton cmdPrint 
            BackColor       =   &H80000009&
            Caption         =   "x4"
            Height          =   270
            Index           =   4
            Left            =   90
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   1065
            Width           =   555
         End
         Begin VB.CommandButton cmdPrint 
            BackColor       =   &H80000009&
            Caption         =   "x3"
            Height          =   270
            Index           =   3
            Left            =   90
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   780
            Width           =   555
         End
         Begin VB.CommandButton cmdPrint 
            BackColor       =   &H80000009&
            Caption         =   "x2"
            Height          =   270
            Index           =   2
            Left            =   90
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   495
            Width           =   555
         End
         Begin VB.CommandButton cmdPrint 
            BackColor       =   &H80000009&
            Caption         =   "x1"
            Height          =   270
            Index           =   1
            Left            =   90
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   210
            Width           =   555
         End
      End
      Begin VB.PictureBox pic1 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   3000
         Left            =   150
         Picture         =   "Main.frx":14CE
         ScaleHeight     =   200
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   180
         TabIndex        =   7
         Top             =   150
         Width           =   2700
      End
      Begin VB.Label Label1 
         BackColor       =   &H00008000&
         Caption         =   "Extractor  by  Robert Rayment"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   2070
         Left            =   105
         TabIndex        =   15
         Top             =   3510
         Width           =   3900
      End
   End
   Begin VB.Frame fraScr 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   1
      Left            =   540
      TabIndex        =   1
      Top             =   6570
      Width           =   3030
      Begin VB.CommandButton cmdScr 
         BackColor       =   &H0000C000&
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   2745
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   180
         Width           =   255
      End
      Begin VB.CommandButton cmdScr 
         BackColor       =   &H0000C000&
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   180
         Width           =   255
      End
      Begin VB.Label LabThumb 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   1
         Left            =   360
         TabIndex        =   13
         Top             =   180
         Width           =   210
      End
      Begin VB.Label LabBack 
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   330
         TabIndex        =   12
         Top             =   165
         Width           =   2400
      End
   End
   Begin VB.Frame fraScr 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Height          =   3150
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   780
      Width           =   375
      Begin VB.CommandButton cmdScr 
         BackColor       =   &H0000C000&
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Index           =   1
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2865
         Width           =   270
      End
      Begin VB.CommandButton cmdScr 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Index           =   0
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   195
         Width           =   270
      End
      Begin VB.Label LabThumb 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   0
         Left            =   90
         TabIndex        =   3
         Top             =   1020
         Width           =   210
      End
      Begin VB.Label LabBack 
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Height          =   3975
         Index           =   0
         Left            =   75
         TabIndex        =   2
         Top             =   -660
         Width           =   255
      End
   End
   Begin VB.CommandButton cmdShiftFocus 
      Caption         =   "Command1"
      Height          =   195
      Left            =   255
      TabIndex        =   50
      Top             =   7110
      Width           =   120
   End
   Begin VB.Label LabpicL 
      BackColor       =   &H00008000&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   3
      Left            =   6105
      TabIndex        =   45
      Top             =   120
      Width           =   435
   End
   Begin VB.Label LabpicL 
      BackColor       =   &H00008000&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   2
      Left            =   4530
      TabIndex        =   44
      Top             =   120
      Width           =   435
   End
   Begin VB.Label LabpicL 
      BackColor       =   &H00008000&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   1
      Left            =   2940
      TabIndex        =   43
      Top             =   90
      Width           =   435
   End
   Begin VB.Label LabpicL 
      BackColor       =   &H00008000&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   0
      Left            =   7875
      TabIndex        =   42
      Top             =   120
      Width           =   435
   End
   Begin VB.Label LabRGBSel 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   4875
      TabIndex        =   41
      Top             =   420
      Width           =   360
   End
   Begin VB.Label LabRGBSel 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   4515
      TabIndex        =   40
      Top             =   420
      Width           =   360
   End
   Begin VB.Label LabRGBSel 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   4155
      TabIndex        =   39
      Top             =   420
      Width           =   360
   End
   Begin VB.Label LabRGBSel 
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Index           =   0
      Left            =   3675
      TabIndex        =   38
      ToolTipText     =   "Selected color"
      Top             =   420
      Width           =   480
   End
   Begin VB.Label LabRGBVary 
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Index           =   0
      Left            =   1800
      TabIndex        =   34
      Top             =   420
      Width           =   480
   End
   Begin VB.Label LabRGBVary 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   2985
      TabIndex        =   33
      Top             =   420
      Width           =   360
   End
   Begin VB.Label LabRGBVary 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   2625
      TabIndex        =   32
      Top             =   420
      Width           =   360
   End
   Begin VB.Label LabRGBVary 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   2265
      TabIndex        =   31
      Top             =   420
      Width           =   360
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   7275
      Left            =   -8835
      Top             =   -6690
      Width           =   9330
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Extractor by  Robert Rayment  Dec 2002

' Extracts a band of colors from a picture and
' shows them on a black & white picture which
' can be saved as a 1-bpp bmp mask.
' Also can extract rectangles or shapes plus a mask
' or resize to a rectangle.
' ASM can be toggled for large pictures.

' NB For shape drawing Shift & Ctrl keys force straight lines.

Option Base 1

DefLng A-W  ' Longs
DefSng X-Z  ' Singles

'Public pic1mem() As Byte   ' For Loaded picture
'Public pic2mem() As Byte   ' For Output picture

Dim picsave() As Byte   ' For saving 2-color picture

Dim zHSc, zVSc    ' Resizing factors
Dim FW, FH        ' Reduced Form width & height


' For using OSDialog(FileDlg2.cls)
Dim CommonDialog1 As New OSDialog

Dim Pathspec$, CurrPath$


Private Sub Form_Load()

If App.PrevInstance Then End

ClearFrames

pic1.ScaleMode = vbpixel
pic2.ScaleMode = vbpixel

Pathspec$ = App.Path
If Right$(Pathspec$, 1) <> "\" Then Pathspec$ = Pathspec$ & "\"
CurrPath$ = Pathspec$

pic1.Move 0, 0

'pic1.Picture = LoadPicture(Pathspec$ & "MandelCastle.bmp")

picHt = pic1.Height
picWd = pic1.Width

pic2.Move 0, 0, pic1.Width, pic1.Height

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
GETDIBS pic1.Image  ' BPP Fills pic1mem and sizes pic2mem

'------------------------------
' Scroll bars set up in calling Form:-
'Set frm = Form1
ReDim scrMax(0 To 1)
ReDim scrMin(0 To 1)
ReDim SmallChange(0 To 1)
ReDim LargeChange(0 To 1)
'------------------------------

RESIZER

cmdPrint(0).Enabled = False
cmdSavePic(0).Enabled = False
cmdSavePic(1).Enabled = False
ASMVB = False

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Load Machine code from bin file
'Loadmcode Pathspec$ & "Extract.bin", ExtractMC()
'ptrMC = VarPtr(ExtractMC(1))
'ptrStruc = VarPtr(DStruc.picWd)

'Load Machine code frm Res file
'Public ExtractMC() As Byte  'Array to hold machine code
'Public ptrMC, ptrStruc      ' Ptrs to Machine Code & Structure
ExtractMC = LoadResData("EXTRACT", "ASM")
ptrMC = VarPtr(ExtractMC(0))
ptrStruc = VarPtr(DStruc.picWd)

'BB = ExtractMC(0)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
DStruc.picHt = picHt
DStruc.picWd = picWd
DStruc.Ptrpic1mem = VarPtr(pic1mem(1, 1, 1))
DStruc.Ptrpic2mem = VarPtr(pic2mem(1, 1, 1))
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

' Default +/- ranges
ReDim DELRGB(0 To 3)
DELRGB(0) = 10
DELRGB(1) = 10
DELRGB(2) = 10
DELRGB(3) = 10
For i = 0 To 3
   LabpicL(i).Caption = Chr$(177) & DELRGB(0)
Next i

A$ = "Extractor  by  Robert Rayment" & vbLf & vbLf
A$ = A$ & " Move mouse over picture and click" & vbLf
A$ = A$ & " to select a color.  Adjust ranges" & vbLf
A$ = A$ & " for all on grey bar and individuals" & vbLf
A$ = A$ & " on R, G or B bars." & vbLf & vbLf
A$ = A$ & " Tools menu for extractors" & vbLf

Label1.Caption = A$

cmdTest.Enabled = False

Show
End Sub

Private Sub pic1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LngToRGB GetPixel(pic1.hdc, X, Y)
LabRGBVary(0).BackColor = RGB(red, green, blue)
LabRGBVary(1).Caption = red
LabRGBVary(2).Caption = green
LabRGBVary(3).Caption = blue
End Sub


'#### TRANSFER SUBS #####################################################

Private Sub pic1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
cmdTest.Enabled = True

Cul = GetPixel(pic1.hdc, X, Y)
If Cul <> -1 Then StartColor = Cul Else Exit Sub

LngToRGB StartColor

LabRGBSel(0).BackColor = RGB(red, green, blue)
LabRGBSel(1).Caption = red
LabRGBSel(2).Caption = green
LabRGBSel(3).Caption = blue

TRANSFER Form1

End Sub

Private Sub picLims_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

cmdTest.Enabled = True

Select Case Index
Case 0   ' Grey DEL all
   DELRGB(0) = 100 * X / picLims(0).Width
   DELRGB(1) = DELRGB(0)   ' RED
   DELRGB(2) = DELRGB(0)   ' GREEN
   DELRGB(3) = DELRGB(0)   ' BLUE
   For i = 0 To 3
      LabpicL(i).Caption = Chr$(177) & DELRGB(0)
   Next i
Case Else
   DELRGB(Index) = 100 * X / picLims(Index).Width
   LabpicL(Index).Caption = Chr$(177) & DELRGB(Index)
End Select

TRANSFER Form1

End Sub

Private Sub picLims_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
Case 0   ' Grey DEL all
   DELRGB(0) = 100 * X / picLims(0).Width
   DELRGB(1) = DELRGB(0)   ' RED
   DELRGB(2) = DELRGB(0)   ' GREEN
   DELRGB(3) = DELRGB(0)   ' BLUE
   For i = 0 To 3
      LabpicL(i).Caption = Chr$(177) & DELRGB(0)
   Next i
Case Else
   DELRGB(Index) = 100 * X / picLims(Index).Width
   LabpicL(Index).Caption = Chr$(177) & DELRGB(Index)
End Select

TRANSFER Form1

End Sub

Private Sub cmdSetPM_Click(Index As Integer)
Select Case Index
Case 0: DELRGB(0) = 0   ' +/- 0
Case 1: DELRGB(0) = 10  ' +/- 10
Case 2: DELRGB(0) = 20  ' +/- 20
Case 3: DELRGB(0) = 30  ' +/- 30
Case 4: DELRGB(0) = 50  ' +/- 50
Case 5: DELRGB(0) = 100 ' +/- 100
End Select

DELRGB(1) = DELRGB(0)   ' RED
DELRGB(2) = DELRGB(0)   ' GREEN
DELRGB(3) = DELRGB(0)   ' BLUE
For i = 0 To 3
   LabpicL(i).Caption = Chr$(177) & DELRGB(0)
Next i

TRANSFER Form1

End Sub

'#### TOOLS #################################################

Private Sub cmdInvert_Click()
res = SetRect(IR, 0, 0, picWd, picHt)
res = InvertRect(pic2.hdc, IR)
pic2.Refresh

' Invert for saving
bm.Colors(0).rgbBlue = 255 - bm.Colors(0).rgbBlue
bm.Colors(0).rgbGreen = 255 - bm.Colors(0).rgbGreen
bm.Colors(0).rgbRed = 255 - bm.Colors(0).rgbRed

bm.Colors(1).rgbBlue = 255 - bm.Colors(1).rgbBlue
bm.Colors(1).rgbGreen = 255 - bm.Colors(1).rgbGreen
bm.Colors(1).rgbRed = 255 - bm.Colors(1).rgbRed
End Sub

Private Sub cmdTest_Click()
TestDrawMode = True  ' ie Test Mask mode
Unload Form2
Form2.Show
End Sub

Private Sub cmdDrawMode_Click()
TestDrawMode = False ' ie Draw mode
Unload Form2
Form2.Show
End Sub


'###### LOAD PICTURE ###############################################

Private Sub cmdLoadPic_Click()
Unload Form2
On Error GoTo LoadError

pic1.Enabled = False
DoEvents

Title$ = "Load a picture file"
Filt$ = "Pics bmp,jpg,gif,ico,cur,wmf,emf|*.bmp;*.jpg;*.gif;*.ico;*.cur;*.wmf;*.emf"
InDir$ = CurrPath$ 'Pathspec$

CommonDialog1.ShowOpen FileSpec$, Title$, Filt$, InDir$, ""

If Len(FileSpec$) = 0 Then
   Close
   pic1.Enabled = True
   Exit Sub
End If

pic1.Cls
picWd = 100
picHt = 100
pic1.Width = picWd
pic1.Height = picHt
ReDim pic1mem(4, picWd, picHt)
ReDim pic2mem(4, picWd, picHt)

pic1.Picture = LoadPicture
pic2.Picture = LoadPicture
DoEvents

pic1.Picture = LoadPicture(FileSpec$)
DoEvents

CurrPath$ = FileSpec$

picWd = pic1.Width
picHt = pic1.Height

If picWd > 1152 Or picHt > 1152 Then
   MsgBox "Picture" & Str$(picWd) & "x" & Str$(picHt) & " a bit too big", , "Extractor"
   pic1.Cls
   picWd = 100
   picHt = 100
   pic1.Width = picWd
   pic1.Height = picHt
   ReDim pic1mem(4, picWd, picHt)
   ReDim pic2mem(4, picWd, picHt)
   pic1.Enabled = True
   Exit Sub
End If

pic2.Move 0, 0, pic1.Width, pic1.Height

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
GETDIBS pic1.Image  ' BPP Fills pic1mem and sizes pic2mem

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
DStruc.picWd = picWd
DStruc.picHt = picHt
DStruc.Ptrpic1mem = VarPtr(pic1mem(1, 1, 1))
DStruc.Ptrpic2mem = VarPtr(pic2mem(1, 1, 1))
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Timer1.Enabled = False

fraScr(0).Visible = False
fraScr(1).Visible = False

ClearFrames

RESIZER

pic1.Enabled = True
DoEvents

On Error GoTo 0
Exit Sub
'==========
LoadError:
Erase pic1mem, pic2mem
DoEvents
pic1.Picture = LoadPicture
pic2.Picture = LoadPicture
MsgBox "VB doesn't like this picture ??", vbCritical, "Extractor"
   pic1.Cls
   picWd = 100
   picHt = 100
   pic1.Width = picWd
   pic1.Height = picHt
   ReDim pic1mem(4, picWd, picHt)
   ReDim pic2mem(4, picWd, picHt)
   pic1.Enabled = True
Exit Sub
End Sub

'###### SAVE FILE ##################################################

Private Sub cmdSavePic_Click(Index As Integer)

On Error GoTo SaveError
Select Case Index
Case 0   ' Save Mask

   Title$ = "Save Mask as 2-color bmp"
   Filt$ = "Save bmp|*.bmp"
   InDir$ = CurrPath$ 'Pathspec$
   
   pic1.Enabled = False
   
   CommonDialog1.ShowSave FileSpec$, Title$, Filt$, InDir$, ""
   
   If Len(FileSpec$) = 0 Then
      Close
      pic1.Enabled = True
      Exit Sub
   End If
   
   ' For bpp=1
   BScanLine = (picWd + 7) \ 8
   ' Expand to 4B boundary
   BScanLine = ((BScanLine + 3) \ 4) * 4
   
   ReDim picsave(BScanLine, picHt)  ' BScanLine bytes
   
   ' Transfer pic2mem to 2-color picsave
   ' (1,1,1)(2,1,1)(3,1,1)(4,1,1)... (1,picWd,1)(2,picWd,1)(3,picWd,1)(4,picWd,1)
   '   0 (< 20)or 255(>250)
   ' to (building bits from 4-bytes)
   ' picsave(1,1)....(picWd\8,1)
   '
   ' picsave(1,picHt)....(picWd\8,picHt)
   
   For j = 1 To picHt
      n = 1
      B = 1
      For i = 1 To picWd
         If pic2mem(1, i, j) > 250 Then
            picsave(n, j) = picsave(n, j) Or 1
            If B < 8 Then picsave(n, j) = picsave(n, j) * 2
         Else
            If B < 8 Then picsave(n, j) = picsave(n, j) * 2
         End If
         B = B + 1
         If B = 9 Then
            n = n + 1
            B = 1
         End If
      Next i
      If B <> 1 Then
         picsave(n, j) = picsave(n, j) * 2 ^ (7 - B + 1)
      End If
   Next j
      
   ''''''''''''''''''''''''''''''''''''''''''''''''''
   ' Save 32 bpp & image size
   svbiBitCount = bm.bmiH.biBitCount
   svbiSizeImage = bm.bmiH.biSizeImage
   ''''''''''''''''''''''''''''''''''''''''''''''''''
   
   bm.bmiH.biBitCount = 1
   bm.bmiH.biSizeImage = BScanLine * picHt
   
   Open FileSpec$ For Binary As #1
   TheBM% = 19778: Put #1, , TheBM%
   FileSize& = 62 + BScanLine * picHt: Put #1, , FileSize&
   ires% = 0: Put #1, , ires%: Put #1, , ires%
   off& = 62: Put #1, , off&
   '' BITMAPINFOHEADER =
   Put #1, , bm
   Put #1, , picsave
   Close
   Erase picsave
   pic1.Enabled = True
   ''''''''''''''''''''''''''''''''''''''''''''''''''
   ' Restore 32 bpp & image size
   bm.bmiH.biBitCount = svbiBitCount
   bm.bmiH.biSizeImage = svbiSizeImage
   ''''''''''''''''''''''''''''''''''''''''''''''''''
   
Case 1   ' Save picture

   Title$ = "Save Picture as 24 bpp"
   Filt$ = "Save bmp|*.bmp"
   InDir$ = CurrPath$ 'Pathspec$
   
   pic1.Enabled = False
   
   CommonDialog1.ShowSave FileSpec$, Title$, Filt$, InDir$, ""
   
   If Len(FileSpec$) = 0 Then
      Close
      pic1.Enabled = True
      Exit Sub
   End If
   
   SavePicture pic1.Image, FileSpec$

   pic1.Enabled = True

End Select

ClearFrames
Exit Sub
'=============
SaveError:
MsgBox "Error in saving", , "Extractor"
Close
Erase picsave
''''''''''''''''''''''''''''''''''''''''''''''''''
' Restore 32 bpp & image size
bm.bmiH.biBitCount = svbiBitCount
bm.bmiH.biSizeImage = svbiSizeImage
''''''''''''''''''''''''''''''''''''''''''''''''''
ClearFrames
pic1.Enabled = True
End Sub

'###### PRINTING ###################################################

Private Sub cmdPrint_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Select Case Index
Case 0
   Frame3.Visible = True
   Exit Sub
Case 1, 2, 3, 4
Case Else
   Exit Sub
End Select

If MsgBox("Print (x" & Str$(Index) & ") to " & Printer.DeviceName, vbYesNo, "Extractor") = vbNo Then
   Frame3.Visible = False
   Exit Sub
End If

' Print Index x real size, offset 400,400 printer units
pic1.Enabled = False

Screen.MousePointer = vbHourglass

NewDC = CreateCompatibleDC(0&)
OldH = SelectObject(NewDC, PICIM)
ret = GetDIBits(NewDC, pic2.Image, 0, picHt, pic2mem(1, 1, 1), bm, DIB_RGB_COLORS)
SelectObject NewDC, OldH
DeleteDC NewDC

Printer.Print " ";

If StretchDIBits(Printer.hdc, 400, 400, Index * picWd, Index * picHt, 0, 0, _
   picWd, picHt, pic2mem(1, 1, 1), bm, DIB_RGB_COLORS, vbSrcCopy) = 0 Then
      MsgBox "Printing failed", , "Extractor"
      Printer.EndDoc
      Frame1.Visible = False
      Frame2.Visible = False
      Frame3.Visible = False
      pic1.Enabled = True
      Screen.MousePointer = vbDefault
      Exit Sub
   End If

Printer.NewPage
Printer.EndDoc

ClearFrames
pic1.Enabled = True
Screen.MousePointer = vbDefault
End Sub

Private Sub cmdVBASM_Click()
ASMVB = Not ASMVB
If ASMVB Then
   cmdVBASM.Caption = "ASM"
   cmdVBASM.BackColor = RGB(255, 200, 0)
Else
   cmdVBASM.Caption = "VB"
   cmdVBASM.BackColor = RGB(255, 255, 255)
End If
DoEvents
End Sub


'###### INITIALIZE #################################################

Private Sub Form_Initialize()

'Dim zHSc, zVSc

' Get screen res
FWW = GetSystemMetrics(SM_CXSCREEN)
FHH = GetSystemMetrics(SM_CYSCREEN)
' To show Desktop border
FW = 0.95 * FWW
FH = 0.9 * FHH
' Scaling for screen res changes
zHSc = (350 / 800) * (FWW / FHH) * 3 / 4
zVSc = (500 / 600) * (FWW / FHH) * 3 / 4

End Sub

'###### RESIZING ###################################################

Private Sub chkResizer_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ClearFrames
chkResizer.Value = Unchecked
RESIZER
End Sub

Private Sub RESIZER()

' Screen res resizer

If WindowState = vbMinimized Then Exit Sub

' Get screen res
FWW = GetSystemMetrics(SM_CXSCREEN)
FHH = GetSystemMetrics(SM_CYSCREEN)
' To show Desktop border
FW = 0.95 * FWW
FH = 0.9 * FHH
 
If WindowState <> vbMaximized Then
   Form1.Top = 250
   Form1.Left = 250
   Form1.Width = FW * Screen.TwipsPerPixelX
   Form1.Height = FH * Screen.TwipsPerPixelY
End If

Picture1.Width = FW * zHSc
Picture1.Height = FH * zVSc

Picture1.Top = 50
pic1.Top = 0

With fraScr(0)
   .Left = 10
   .Top = Picture1.Top + 2
   .Width = 20
   .Height = Picture1.Height
End With

Picture1.Left = fraScr(0).Left + fraScr(0).Width + 5
pic1.Left = 0

With fraScr(1)
   .Top = Picture1.Top + Picture1.Height + 2
   .Left = Picture1.Left
   .Height = 20
   .Width = Picture1.Width
End With

With Picture2
   .Top = Picture1.Top
   .Left = Picture1.Left + Picture1.Width + 10
   .Height = Picture1.Height
   .Width = Picture1.Width
End With
pic2.Top = 0
pic2.Left = 0

' Shape picLims
TLX = 0
TLY = 0
For i = 0 To 3
   BRX = picLims(i).Width
   BRY = picLims(i).Height
   RRRPic = CreateRoundRectRgn _
   (TLX, TLY, BRX, BRY, 20, 20)
   SetWindowRgn picLims(i).hwnd, RRRPic, False
   DeleteObject RRRPic
Next i
   
zincr = picLims(1).Width / 256
Cul = 0
For zpx = 0 To picLims(0).Width Step zincr
   picLims(0).Line (zpx, 0)-(zpx, picLims(0).Height), RGB(Cul, Cul, Cul)
   picLims(1).Line (zpx, 0)-(zpx, picLims(0).Height), RGB(Cul, 0, 0)
   picLims(2).Line (zpx, 0)-(zpx, picLims(0).Height), RGB(0, Cul, 0)
   picLims(3).Line (zpx, 0)-(zpx, picLims(0).Height), RGB(0, 0, Cul)
   Cul = Cul + 1
   If Cul > 255 Then Cul = 255
Next zpx

' Other controls
cmdExit(1).Left = Form1.Width / Screen.TwipsPerPixelX - cmdExit(1).Width - 6
Shape1.Move 3, 3, Form1.Width / Screen.TwipsPerPixelX - 6, Form1.Height / Screen.TwipsPerPixelY - 6

cmdSetPM(0).Caption = Chr$(177) & "0"
cmdSetPM(1).Caption = Chr$(177) & "10"
cmdSetPM(2).Caption = Chr$(177) & "20"
cmdSetPM(3).Caption = Chr$(177) & "30"
cmdSetPM(4).Caption = Chr$(177) & "50"
cmdSetPM(5).Caption = Chr$(177) & "100"

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

If picHt < 4 Or picWd < 4 Then
   MsgBox "Very small or thin picture", , "Extractor"
End If

' SET UP SCROLL BARS

ModReDim Form1
nTVal(0) = scrMin(0)
nTVal(1) = scrMin(1)

cmdShiftFocus.Visible = False ' Shift focus from cmdScr(0,1/2,3)

If picHt > Picture1.Height Then
   fraScr(0).Visible = True
   scrMax(0) = picHt - Picture1.Height + 6
   LargeChange(0) = scrMax(0) \ 10
   If LargeChange(0) < 1 Then LargeChange(0) = 1
Else
   fraScr(0).Visible = False
End If

If picWd > Picture1.Width Then
   fraScr(1).Visible = True
   scrMax(1) = picWd - Picture1.Width + 6
   LargeChange(1) = scrMax(1) \ 10
   If LargeChange(1) < 1 Then LargeChange(1) = 1
Else
   fraScr(1).Visible = False
End If
' Initial scroll bar values
VVal = nTVal(0)
HVal = nTVal(1)
'----------------------------------------------

End Sub

'###### EXIT #######################################################

Private Sub Form_Unload(Cancel As Integer)
Dim Form As Form
Erase pic1mem, pic2mem
' Make sure all forms cleared
For Each Form In Forms
   Unload Form
   Set Form = Nothing
Next Form
End
End Sub

Private Sub cmdExit_Click(Index As Integer)
ClearFrames

Unload Form2
resp = MsgBox("Quit ?", vbQuestion + vbYesNo, "Extractor")
If resp = vbYes Then Form_Unload 1

End Sub

'###### TOGGLE FRAMES ##############################################

Private Sub cmdFile_Click()
Frame1.Visible = False
Frame2.Visible = Not Frame2.Visible
Frame3.Visible = False
End Sub
Private Sub cmdTools_Click()
Frame1.Visible = Not Frame1.Visible
Frame2.Visible = False
Frame3.Visible = False
End Sub

'###### CLEAR FRAMES ###############################################

Private Sub Form_Click()
ClearFrames
End Sub

Private Sub Frame1_Click()
ClearFrames
End Sub

Private Sub Frame2_Click()
ClearFrames
End Sub

Private Sub Frame3_Click()
ClearFrames
End Sub

Private Sub Picture1_Click()
ClearFrames
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ClearFrames
End Sub

Private Sub pic2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ClearFrames
End Sub

Private Sub ClearFrames()
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
End Sub


'###################################################################
'####### SCROLL BARS ###############################################
'###################################################################

Private Sub LabThumb_MouseDown(Index As Integer, Button As Integer, Shift As Integer, _
X As Single, Y As Single)
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
   LabThumbMouseDown Form1, Index, X, Y
   
   If nTVal(0) <> VVal Or nTVal(1) <> HVal Then
      MovePics
      VVal = nTVal(0)
      HVal = nTVal(1)
   End If
   
End Sub

Private Sub LabThumb_MouseMove(Index As Integer, Button As Integer, Shift As Integer, _
X As Single, Y As Single)
   LabThumbMouseMove Form1, Index, X, Y
   
   If nTVal(0) <> VVal Or nTVal(1) <> HVal Then
      MovePics
      VVal = nTVal(0)
      HVal = nTVal(1)
   End If

End Sub

Private Sub LabThumb_MouseUp(Index As Integer, Button As Integer, Shift As Integer, _
X As Single, Y As Single)
   ' Snap thumb position
   LabThumbMouseUp Form1, Index, X, Y
   
   If nTVal(0) <> VVal Or nTVal(1) <> HVal Then
      MovePics
      VVal = nTVal(0)
      HVal = nTVal(1)
   End If

End Sub

'###################################################################

Private Sub LabBack_MouseDown(Index As Integer, Button As Integer, Shift As Integer, _
X As Single, Y As Single)
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
   
   LabBackMouseDown Form1, Index, X, Y
   
   If nTVal(0) <> VVal Or nTVal(1) <> HVal Then
      MovePics
      VVal = nTVal(0)
      HVal = nTVal(1)
   End If

End Sub

Private Sub LabBack_MouseUp(Index As Integer, Button As Integer, Shift As Integer, _
X As Single, Y As Single)
   LabBackMouseUp Form1
   
   If nTVal(0) <> VVal Or nTVal(1) <> HVal Then
      MovePics
      VVal = nTVal(0)
      HVal = nTVal(1)
   End If

End Sub

'###################################################################

Private Sub cmdScr_MouseDown(Index As Integer, Button As Integer, Shift As Integer, _
X As Single, Y As Single)
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
   
   cmdScrMouseDown Form1, Index, Button, X, Y
End Sub

Private Sub cmdScr_MouseUp(Index As Integer, Button As Integer, Shift As Integer, _
X As Single, Y As Single)
Frame2.Visible = False
   
   cmdScrMouseUp Form1
End Sub

Private Sub cmdScr_Click(Index As Integer)

   cmdScrMouseUp Form1

End Sub


'###################################################################

Private Sub Timer1_Timer()
   Timer1Timer Form1
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
pic1.Top = -nTVal(0)
pic1.Refresh
pic2.Top = -nTVal(0)
pic2.Refresh
' Horz
pic1.Left = -nTVal(1)
pic1.Refresh
pic2.Left = -nTVal(1)
pic2.Refresh
End Sub

