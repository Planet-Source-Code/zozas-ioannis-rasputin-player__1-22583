VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   6600
   ClientLeft      =   660
   ClientTop       =   660
   ClientWidth     =   10560
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   161
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   HasDC           =   0   'False
   Icon            =   "fmrMain.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   10560
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0080FFFF&
      Height          =   4230
      ItemData        =   "fmrMain.frx":030A
      Left            =   120
      List            =   "fmrMain.frx":030C
      TabIndex        =   4
      Top             =   2280
      Width           =   3975
   End
   Begin VB.ListBox TempList 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10200
      Sorted          =   -1  'True
      TabIndex        =   29
      Top             =   1680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ListBox ListPresets 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   1785
      ItemData        =   "fmrMain.frx":030E
      Left            =   4320
      List            =   "fmrMain.frx":033C
      Sorted          =   -1  'True
      TabIndex        =   28
      ToolTipText     =   "Equalizer Presets"
      Top             =   4440
      Width           =   2775
   End
   Begin VB.VScrollBar Preamp 
      Height          =   1335
      Index           =   4
      LargeChange     =   10
      Left            =   6840
      Max             =   -50
      Min             =   50
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   4560
      Width           =   135
   End
   Begin VB.VScrollBar Preamp 
      Height          =   1335
      Index           =   3
      LargeChange     =   10
      Left            =   6480
      Max             =   -50
      Min             =   50
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   4560
      Width           =   135
   End
   Begin VB.VScrollBar Preamp 
      Height          =   1335
      Index           =   2
      LargeChange     =   10
      Left            =   6120
      Max             =   -50
      Min             =   50
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   4560
      Width           =   135
   End
   Begin VB.VScrollBar Preamp 
      Height          =   1335
      Index           =   1
      LargeChange     =   10
      Left            =   5760
      Max             =   -50
      Min             =   50
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   4560
      Width           =   135
   End
   Begin VB.VScrollBar Equ 
      Height          =   1335
      Left            =   4440
      TabIndex        =   19
      Top             =   4560
      Value           =   32767
      Width           =   255
   End
   Begin VB.VScrollBar Preamp 
      Height          =   1335
      Index           =   0
      LargeChange     =   10
      Left            =   5400
      Max             =   -50
      Min             =   50
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4560
      Width           =   135
   End
   Begin VB.TextBox TxtName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   330
      Left            =   240
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "Rasputin - No Title *** "
      ToolTipText     =   "Track File"
      Top             =   120
      Width           =   3375
   End
   Begin VB.Timer TimerTitle 
      Interval        =   150
      Left            =   9720
      Top             =   2040
   End
   Begin MSComctlLib.Slider sldVol 
      Height          =   1035
      Left            =   4320
      TabIndex        =   6
      ToolTipText     =   "Song Volume"
      Top             =   1800
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1826
      _Version        =   393216
      BorderStyle     =   1
      Orientation     =   1
      Max             =   2500
      TickStyle       =   2
      TickFrequency   =   125
      TextPosition    =   1
   End
   Begin VB.TextBox TxtTime 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   225
      Left            =   600
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "00:00/00:00"
      Top             =   480
      Width           =   1215
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   300
      Left            =   240
      TabIndex        =   14
      Top             =   6120
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   10200
      Pattern         =   "*.mp3;*.wav"
      TabIndex        =   15
      Top             =   2040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0080FFFF&
      Height          =   3450
      Left            =   120
      TabIndex        =   13
      Top             =   2280
      Visible         =   0   'False
      Width           =   3975
   End
   Begin MSComDlg.CommonDialog CommonDialog3 
      Left            =   9720
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox ListInfo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   225
      Left            =   1920
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "10000/10000 File"
      Top             =   480
      Width           =   1695
   End
   Begin VB.Timer TimerPlayer 
      Interval        =   1
      Left            =   9720
      Top             =   1560
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9720
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   9720
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Slider sldCenter 
      Height          =   1035
      Left            =   5040
      TabIndex        =   7
      ToolTipText     =   "Centralization"
      Top             =   1800
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1826
      _Version        =   393216
      BorderStyle     =   1
      Orientation     =   1
      LargeChange     =   1
      Min             =   -5000
      Max             =   5000
      TickStyle       =   2
      TickFrequency   =   500
      TextPosition    =   1
   End
   Begin MSComctlLib.Slider sldSearch 
      Height          =   615
      Left            =   4320
      TabIndex        =   5
      ToolTipText     =   "Scanning Bar"
      Top             =   3240
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1085
      _Version        =   393216
      BorderStyle     =   1
      LargeChange     =   2
      Max             =   100
      TickStyle       =   2
      TickFrequency   =   10
      TextPosition    =   1
   End
   Begin ComctlLib.Slider sldMainBas 
      Height          =   1035
      Left            =   6480
      TabIndex        =   9
      ToolTipText     =   "Bass Level"
      Top             =   1800
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1826
      _Version        =   327682
      BorderStyle     =   1
      Orientation     =   1
      TickStyle       =   2
   End
   Begin ComctlLib.Slider sldMainTreb 
      Height          =   1035
      Left            =   5760
      TabIndex        =   8
      ToolTipText     =   "Treble Level"
      Top             =   1800
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1826
      _Version        =   327682
      BorderStyle     =   1
      Orientation     =   1
      TickStyle       =   2
   End
   Begin VB.Image ImgCUtilityon 
      Height          =   225
      Left            =   8760
      Picture         =   "fmrMain.frx":03B2
      Top             =   1920
      Width           =   660
   End
   Begin VB.Image ImgCUtilityoff 
      Height          =   225
      Left            =   8760
      Picture         =   "fmrMain.frx":0BB0
      Top             =   1680
      Width           =   660
   End
   Begin VB.Image cmdCUtility 
      Height          =   225
      Left            =   840
      Picture         =   "fmrMain.frx":13AE
      ToolTipText     =   "Add sound file collections (directories and subdirectories)"
      Top             =   5880
      Width           =   660
   End
   Begin VB.Image ImgScanon 
      Height          =   225
      Left            =   8760
      Picture         =   "fmrMain.frx":1BAC
      Top             =   2400
      Width           =   660
   End
   Begin VB.Image ImgScanoff 
      Height          =   225
      Left            =   8760
      Picture         =   "fmrMain.frx":23AA
      Top             =   2160
      Width           =   660
   End
   Begin VB.Image ImgBack_on 
      Height          =   225
      Left            =   8760
      Picture         =   "fmrMain.frx":2BA8
      Top             =   2880
      Width           =   660
   End
   Begin VB.Image ImgBack_Off 
      Height          =   225
      Left            =   8760
      Picture         =   "fmrMain.frx":33A6
      Top             =   2640
      Width           =   660
   End
   Begin VB.Image ImgScanSubs_on 
      Height          =   225
      Left            =   8760
      Picture         =   "fmrMain.frx":3BA4
      Top             =   3120
      Width           =   660
   End
   Begin VB.Image ImgScanSubs_off 
      Height          =   225
      Left            =   8760
      Picture         =   "fmrMain.frx":43A2
      Top             =   3360
      Width           =   660
   End
   Begin VB.Image Image_GrEq 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   5520
      Picture         =   "fmrMain.frx":4BA0
      ToolTipText     =   "Show Graphic Equalizer"
      Top             =   600
      Width           =   480
   End
   Begin VB.Label lbl_FI_FR 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0 Frames"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   4320
      TabIndex        =   34
      Top             =   4020
      Width           =   2775
   End
   Begin VB.Label lbl_FI_SR 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "0 Khz"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   180
      Left            =   2910
      TabIndex        =   33
      Top             =   720
      Width           =   345
   End
   Begin VB.Label lbl_FI_BR 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "0 Kbps"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   180
      Left            =   3510
      TabIndex        =   32
      Top             =   720
      Width           =   435
   End
   Begin VB.Label lbl_FI_ID 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "MPEG-X Layer III"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   180
      Left            =   240
      TabIndex        =   31
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label lbl_FI_SZ 
      BackColor       =   &H00000000&
      Caption         =   "00,000 Mb"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   1440
      TabIndex        =   30
      Top             =   720
      Width           =   765
   End
   Begin VB.Image ImgApplyoff 
      Height          =   225
      Left            =   8760
      Picture         =   "fmrMain.frx":4CF2
      Top             =   3600
      Width           =   660
   End
   Begin VB.Image ImgApplyOn 
      Height          =   225
      Left            =   8760
      Picture         =   "fmrMain.frx":54F0
      Top             =   3840
      Width           =   660
   End
   Begin VB.Image ImgOptDefoff 
      Height          =   225
      Left            =   8760
      Picture         =   "fmrMain.frx":5CEE
      Top             =   4080
      Width           =   660
   End
   Begin VB.Image ImgOptDefOn 
      Height          =   225
      Left            =   8760
      Picture         =   "fmrMain.frx":64EC
      Top             =   4320
      Width           =   660
   End
   Begin VB.Label lblPreAmp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Equalizer T/5B T/2B   T/B  2T/B 5T/B"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   4440
      TabIndex        =   20
      Top             =   5880
      Width           =   2655
   End
   Begin VB.Image ImgEQpresetson 
      Height          =   225
      Left            =   8760
      Picture         =   "fmrMain.frx":6CEA
      Top             =   4800
      Width           =   660
   End
   Begin VB.Image ImgEQpresetsoff 
      Height          =   225
      Left            =   8760
      Picture         =   "fmrMain.frx":74E8
      Top             =   4560
      Width           =   660
   End
   Begin VB.Image cmdAmpPreset 
      Height          =   225
      Left            =   5640
      Picture         =   "fmrMain.frx":7CE6
      ToolTipText     =   "Choose an Equalizer Preset Setting"
      Top             =   6240
      Width           =   660
   End
   Begin VB.Image ImgEQloadoff 
      Height          =   225
      Left            =   8760
      Picture         =   "fmrMain.frx":84E4
      Top             =   5040
      Width           =   660
   End
   Begin VB.Image ImgEQloadon 
      Height          =   225
      Left            =   8760
      Picture         =   "fmrMain.frx":8CE2
      Top             =   5280
      Width           =   660
   End
   Begin VB.Image cmdAmpLoad 
      Height          =   225
      Left            =   5040
      Picture         =   "fmrMain.frx":94E0
      ToolTipText     =   "Load Equalizer Settings"
      Top             =   6240
      Width           =   660
   End
   Begin VB.Image ImgEqSaveon 
      Height          =   225
      Left            =   8760
      Picture         =   "fmrMain.frx":9CDE
      Top             =   5760
      Width           =   660
   End
   Begin VB.Image ImgEqSaveoff 
      Height          =   225
      Left            =   8760
      Picture         =   "fmrMain.frx":A4DC
      Top             =   5520
      Width           =   660
   End
   Begin VB.Image cmdAmpSave 
      Height          =   225
      Left            =   4440
      Picture         =   "fmrMain.frx":ACDA
      ToolTipText     =   "Save Equalizer Settings"
      Top             =   6240
      Width           =   660
   End
   Begin VB.Image ImgEQdefoff 
      Height          =   225
      Left            =   8760
      Picture         =   "fmrMain.frx":B4D8
      Top             =   6000
      Width           =   660
   End
   Begin VB.Image ImgEQdefon 
      Height          =   225
      Left            =   8760
      Picture         =   "fmrMain.frx":BCD6
      Top             =   6240
      Width           =   660
   End
   Begin VB.Image cmdAmpDefault 
      Height          =   225
      Left            =   6240
      Picture         =   "fmrMain.frx":C4D4
      ToolTipText     =   "Set Equalizer to Default State"
      Top             =   6240
      Width           =   660
   End
   Begin VB.Label lblm50 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "-50ch"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   255
      Left            =   4845
      TabIndex        =   27
      Top             =   5475
      Width           =   495
   End
   Begin VB.Label lblp50 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "+50ch"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   4800
      TabIndex        =   26
      Top             =   4755
      Width           =   495
   End
   Begin VB.Label lbl0db 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0 ch"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   255
      Left            =   4920
      TabIndex        =   21
      Top             =   5130
      Width           =   375
   End
   Begin VB.Image ImgVoid 
      Height          =   300
      Left            =   9240
      Picture         =   "fmrMain.frx":CCD2
      Top             =   1200
      Width           =   345
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      FillColor       =   &H0080FFFF&
      FillStyle       =   2  'Horizontal Line
      Height          =   855
      Left            =   5280
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Image ImgControlsoff 
      Height          =   300
      Left            =   7920
      Picture         =   "fmrMain.frx":D2B4
      Top             =   120
      Width           =   345
   End
   Begin VB.Image ImgControlson 
      Height          =   300
      Left            =   7920
      Picture         =   "fmrMain.frx":D896
      Top             =   480
      Width           =   345
   End
   Begin VB.Image cmdControls 
      Height          =   300
      Left            =   1920
      Picture         =   "fmrMain.frx":DE78
      Stretch         =   -1  'True
      ToolTipText     =   "Show Controls / Options"
      Top             =   960
      Width           =   345
   End
   Begin VB.Image ImgList 
      Height          =   240
      Left            =   10200
      Picture         =   "fmrMain.frx":E45A
      Stretch         =   -1  'True
      Top             =   480
      Width           =   225
   End
   Begin VB.Image ImgFiles 
      Height          =   240
      Left            =   3720
      Picture         =   "fmrMain.frx":E55C
      Stretch         =   -1  'True
      ToolTipText     =   "Track Number/Total Tracks"
      Top             =   480
      Width           =   225
   End
   Begin VB.Image ImgClock 
      Height          =   225
      Left            =   10200
      Picture         =   "fmrMain.frx":E65E
      Stretch         =   -1  'True
      Top             =   240
      Width           =   225
   End
   Begin VB.Image ImgTime 
      Height          =   225
      Left            =   240
      Picture         =   "fmrMain.frx":E758
      Stretch         =   -1  'True
      ToolTipText     =   "Time Elapsed/Time Remaing"
      Top             =   480
      Width           =   225
   End
   Begin VB.Image ImgAction 
      Height          =   195
      Left            =   3760
      Picture         =   "fmrMain.frx":E852
      Top             =   180
      Width           =   255
   End
   Begin VB.Image ImgMaxoff 
      Height          =   300
      Left            =   7920
      Picture         =   "fmrMain.frx":EB38
      Top             =   840
      Width           =   345
   End
   Begin VB.Image cmdCompact 
      Height          =   300
      Left            =   2280
      Picture         =   "fmrMain.frx":F11A
      Stretch         =   -1  'True
      ToolTipText     =   "Compact mode / List mode"
      Top             =   960
      Width           =   345
   End
   Begin VB.Image ImgMaxon 
      Height          =   300
      Left            =   7920
      Picture         =   "fmrMain.frx":F6FC
      Top             =   1200
      Width           =   345
   End
   Begin VB.Image ImgDeskoff 
      Height          =   225
      Left            =   9480
      Picture         =   "fmrMain.frx":FCDE
      Top             =   2640
      Width           =   660
   End
   Begin VB.Image ImgDeskon 
      Height          =   225
      Left            =   9480
      Picture         =   "fmrMain.frx":104DC
      Top             =   2880
      Width           =   660
   End
   Begin VB.Image cmdDesktop 
      Height          =   225
      Left            =   2040
      Picture         =   "fmrMain.frx":10CDA
      Stretch         =   -1  'True
      ToolTipText     =   "Go to desktop"
      Top             =   5880
      Width           =   660
   End
   Begin VB.Image ImgNull 
      Height          =   195
      Left            =   10200
      Picture         =   "fmrMain.frx":114D8
      Top             =   1200
      Width           =   255
   End
   Begin VB.Image ImgPauseS 
      Height          =   210
      Left            =   10200
      Picture         =   "fmrMain.frx":117BE
      Top             =   1440
      Width           =   255
   End
   Begin VB.Image ImgPlayS 
      Height          =   195
      Left            =   10200
      Picture         =   "fmrMain.frx":11AD8
      Top             =   960
      Width           =   255
   End
   Begin VB.Image ImgStopS 
      Height          =   210
      Left            =   10200
      Picture         =   "fmrMain.frx":11DBE
      Top             =   720
      Width           =   270
   End
   Begin VB.Image ImgMyon 
      Height          =   225
      Left            =   9480
      Picture         =   "fmrMain.frx":12110
      Top             =   3360
      Width           =   660
   End
   Begin VB.Image ImgMyoff 
      Height          =   225
      Left            =   9480
      Picture         =   "fmrMain.frx":1290E
      Top             =   3120
      Width           =   660
   End
   Begin VB.Image cmdMyDocuments 
      Height          =   225
      Left            =   1440
      Picture         =   "fmrMain.frx":1310C
      Stretch         =   -1  'True
      ToolTipText     =   "Go to my documents"
      Top             =   5880
      Width           =   660
   End
   Begin VB.Image ImgLupon 
      Height          =   225
      Left            =   9480
      Picture         =   "fmrMain.frx":1390A
      Top             =   3840
      Width           =   660
   End
   Begin VB.Image cmdLup 
      Height          =   225
      Left            =   2640
      Picture         =   "fmrMain.frx":14108
      Stretch         =   -1  'True
      ToolTipText     =   "Directory level up"
      Top             =   5880
      Width           =   660
   End
   Begin VB.Image ImgLupoff 
      Height          =   225
      Left            =   9480
      Picture         =   "fmrMain.frx":14906
      Top             =   3600
      Width           =   660
   End
   Begin VB.Image ImgCancelon 
      Height          =   225
      Left            =   9480
      Picture         =   "fmrMain.frx":15104
      Top             =   6240
      Width           =   660
   End
   Begin VB.Image ImgCanceloff 
      Height          =   225
      Left            =   9480
      Picture         =   "fmrMain.frx":15902
      Top             =   6000
      Width           =   660
   End
   Begin VB.Image cmdCancel 
      Height          =   225
      Left            =   3240
      Picture         =   "fmrMain.frx":16100
      Stretch         =   -1  'True
      ToolTipText     =   "Cancel addittion"
      Top             =   5880
      Width           =   660
   End
   Begin VB.Image ImgConfirmon 
      Height          =   225
      Left            =   9480
      Picture         =   "fmrMain.frx":168FE
      Top             =   5760
      Width           =   660
   End
   Begin VB.Image ImgConfirmoff 
      Height          =   225
      Left            =   9480
      Picture         =   "fmrMain.frx":170FC
      Top             =   5520
      Width           =   660
   End
   Begin VB.Image cmdOK 
      Height          =   225
      Left            =   240
      Picture         =   "fmrMain.frx":178FA
      Stretch         =   -1  'True
      ToolTipText     =   "Confirm addittion"
      Top             =   5880
      Width           =   660
   End
   Begin VB.Label lblBytes 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Time Elapsed / Total Time"
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   4320
      TabIndex        =   18
      ToolTipText     =   "Total Duration Of Track"
      Top             =   3840
      UseMnemonic     =   0   'False
      Width           =   2775
   End
   Begin VB.Image ImgCenter2off 
      Height          =   225
      Left            =   9480
      Picture         =   "fmrMain.frx":180F8
      Top             =   5280
      Width           =   660
   End
   Begin VB.Image ImgCenter2on 
      Height          =   225
      Left            =   9480
      Picture         =   "fmrMain.frx":188F6
      Top             =   5040
      Width           =   660
   End
   Begin VB.Image cmdTreC 
      Height          =   225
      Left            =   5760
      Picture         =   "fmrMain.frx":190F4
      Stretch         =   -1  'True
      ToolTipText     =   "Center Treble"
      Top             =   2880
      Width           =   660
   End
   Begin VB.Image cmdBassC 
      Height          =   225
      Left            =   6480
      Picture         =   "fmrMain.frx":198F2
      Stretch         =   -1  'True
      ToolTipText     =   "Center Bass"
      Top             =   2880
      Width           =   630
   End
   Begin VB.Label lblTreble 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Treble"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   5760
      TabIndex        =   17
      Top             =   1560
      Width           =   630
   End
   Begin VB.Label lblBass 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Bass"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   6480
      TabIndex        =   16
      Top             =   1560
      Width           =   630
   End
   Begin VB.Image ImgToolson 
      Height          =   300
      Left            =   7920
      Picture         =   "fmrMain.frx":1A0F0
      Top             =   4800
      Width           =   345
   End
   Begin VB.Image cmdTools 
      Height          =   300
      Left            =   1560
      Picture         =   "fmrMain.frx":1A6D2
      Stretch         =   -1  'True
      ToolTipText     =   "Show / Hide Tools"
      Top             =   960
      Width           =   345
   End
   Begin VB.Image ImgToolsoff 
      Height          =   300
      Left            =   7920
      Picture         =   "fmrMain.frx":1ACB4
      Top             =   4440
      Width           =   345
   End
   Begin VB.Image cmdCenter 
      Height          =   225
      Left            =   5040
      Picture         =   "fmrMain.frx":1B296
      Stretch         =   -1  'True
      ToolTipText     =   "Center Volume"
      Top             =   2880
      Width           =   630
   End
   Begin VB.Image ImgCenteron 
      Height          =   225
      Left            =   9480
      Picture         =   "fmrMain.frx":1BA94
      Top             =   4800
      Width           =   660
   End
   Begin VB.Image ImgCenteroff 
      Height          =   225
      Left            =   9480
      Picture         =   "fmrMain.frx":1C292
      Top             =   4560
      Width           =   660
   End
   Begin VB.Image cmdMute 
      Height          =   225
      Left            =   4320
      Picture         =   "fmrMain.frx":1CA90
      Stretch         =   -1  'True
      ToolTipText     =   "Song Mute"
      Top             =   2880
      Width           =   660
   End
   Begin VB.Image ImgMuteon 
      Height          =   225
      Left            =   9480
      Picture         =   "fmrMain.frx":1D28E
      Top             =   4320
      Width           =   660
   End
   Begin VB.Image ImgMuteoff 
      Height          =   225
      Left            =   9480
      Picture         =   "fmrMain.frx":1DA8C
      Top             =   4080
      Width           =   660
   End
   Begin VB.Image cmdOnTop 
      Height          =   300
      Left            =   3360
      Picture         =   "fmrMain.frx":1E28A
      Stretch         =   -1  'True
      ToolTipText     =   "Stay On Top"
      Top             =   960
      Width           =   345
   End
   Begin VB.Image ImgTopon 
      Height          =   300
      Left            =   7920
      Picture         =   "fmrMain.frx":1E86C
      Top             =   2640
      Width           =   345
   End
   Begin VB.Image ImgTopoff 
      Height          =   300
      Left            =   7920
      Picture         =   "fmrMain.frx":1EE4E
      Top             =   2280
      Width           =   345
   End
   Begin VB.Image ImgNoneListon 
      Height          =   300
      Left            =   8280
      Picture         =   "fmrMain.frx":1F430
      Top             =   480
      Width           =   345
   End
   Begin VB.Image ImgNoneListoff 
      Height          =   300
      Left            =   8280
      Picture         =   "fmrMain.frx":1FA12
      Top             =   120
      Width           =   345
   End
   Begin VB.Image cmdNoneList 
      Height          =   300
      Left            =   1200
      Picture         =   "fmrMain.frx":1FFF4
      Stretch         =   -1  'True
      ToolTipText     =   "Don't Play Order"
      Top             =   1920
      Width           =   345
   End
   Begin VB.Image cmdRndList 
      Height          =   300
      Left            =   840
      Picture         =   "fmrMain.frx":205D6
      Stretch         =   -1  'True
      ToolTipText     =   "Play Random Order"
      Top             =   1920
      Width           =   345
   End
   Begin VB.Image ImgRndListon 
      Height          =   300
      Left            =   8280
      Picture         =   "fmrMain.frx":20BB8
      Top             =   1200
      Width           =   345
   End
   Begin VB.Image ImgRndListoff 
      Height          =   300
      Left            =   8280
      Picture         =   "fmrMain.frx":2119A
      Top             =   840
      Width           =   345
   End
   Begin VB.Image ImgNextListon 
      Height          =   300
      Left            =   8280
      Picture         =   "fmrMain.frx":2177C
      Top             =   1920
      Width           =   345
   End
   Begin VB.Image ImgNextListoff 
      Height          =   300
      Left            =   8280
      Picture         =   "fmrMain.frx":21D5E
      Top             =   1560
      Width           =   345
   End
   Begin VB.Image cmdNextList 
      Height          =   300
      Left            =   480
      Picture         =   "fmrMain.frx":22340
      Stretch         =   -1  'True
      ToolTipText     =   "Play Next Order"
      Top             =   1920
      Width           =   345
   End
   Begin VB.Image ImgMinon 
      Height          =   300
      Left            =   8280
      Picture         =   "fmrMain.frx":22922
      Top             =   2640
      Width           =   345
   End
   Begin VB.Image ImgMinoff 
      Height          =   300
      Left            =   8280
      Picture         =   "fmrMain.frx":22F04
      Top             =   2280
      Width           =   345
   End
   Begin VB.Image cmdMinimize 
      Height          =   300
      Left            =   2640
      Picture         =   "fmrMain.frx":234E6
      Stretch         =   -1  'True
      ToolTipText     =   "Minimize"
      Top             =   960
      Width           =   345
   End
   Begin VB.Image ImgLoadLoff 
      Height          =   300
      Left            =   8280
      Picture         =   "fmrMain.frx":23AC8
      Top             =   3000
      Width           =   345
   End
   Begin VB.Image CmdLoadList 
      Height          =   300
      Left            =   1920
      Picture         =   "fmrMain.frx":240AA
      Stretch         =   -1  'True
      ToolTipText     =   "Load PlayList"
      Top             =   1920
      Width           =   345
   End
   Begin VB.Image ImgLoadLon 
      Height          =   300
      Left            =   8280
      Picture         =   "fmrMain.frx":2468C
      Top             =   3360
      Width           =   345
   End
   Begin VB.Image CmdSaveList 
      Height          =   300
      Left            =   1560
      Picture         =   "fmrMain.frx":24C6E
      Stretch         =   -1  'True
      ToolTipText     =   "Save PlayList"
      Top             =   1920
      Width           =   345
   End
   Begin VB.Image ImgSaveLon 
      Height          =   300
      Left            =   8280
      Picture         =   "fmrMain.frx":25250
      Top             =   4080
      Width           =   345
   End
   Begin VB.Image ImgSaveLoff 
      Height          =   300
      Left            =   8280
      Picture         =   "fmrMain.frx":25832
      Top             =   3720
      Width           =   345
   End
   Begin VB.Image cmdClear 
      Height          =   300
      Left            =   3360
      Picture         =   "fmrMain.frx":25E14
      Stretch         =   -1  'True
      ToolTipText     =   "Clear PlayList"
      Top             =   1920
      Width           =   345
   End
   Begin VB.Image ImgRemAon 
      Height          =   300
      Left            =   8280
      Picture         =   "fmrMain.frx":263F6
      Top             =   4800
      Width           =   345
   End
   Begin VB.Image ImgRemAoff 
      Height          =   300
      Left            =   8280
      Picture         =   "fmrMain.frx":269D8
      Top             =   4440
      Width           =   345
   End
   Begin VB.Image cmdRem 
      Height          =   300
      Left            =   3000
      Picture         =   "fmrMain.frx":26FBA
      Stretch         =   -1  'True
      ToolTipText     =   "Remove Track From PlayList"
      Top             =   1920
      Width           =   345
   End
   Begin VB.Image ImgRem1on 
      Height          =   300
      Left            =   8280
      Picture         =   "fmrMain.frx":2759C
      Top             =   5520
      Width           =   345
   End
   Begin VB.Image ImgRem1off 
      Height          =   300
      Left            =   8280
      Picture         =   "fmrMain.frx":27B7E
      Top             =   5160
      Width           =   345
   End
   Begin VB.Image cmdFF 
      Height          =   300
      Left            =   3720
      Picture         =   "fmrMain.frx":28160
      Stretch         =   -1  'True
      ToolTipText     =   "Fast Forward"
      Top             =   1920
      Width           =   365
   End
   Begin VB.Image cmdAdddir 
      Height          =   300
      Left            =   2640
      Picture         =   "fmrMain.frx":28792
      Stretch         =   -1  'True
      ToolTipText     =   "Add Directory To PlayList"
      Top             =   1920
      Width           =   345
   End
   Begin VB.Image ImgAddDon 
      Height          =   300
      Left            =   8280
      Picture         =   "fmrMain.frx":28D74
      Top             =   6240
      Width           =   345
   End
   Begin VB.Image ImgAddDoff 
      Height          =   300
      Left            =   8280
      Picture         =   "fmrMain.frx":29356
      Top             =   5880
      Width           =   345
   End
   Begin VB.Image cmdAddFile 
      Height          =   300
      Left            =   2280
      Picture         =   "fmrMain.frx":29938
      Stretch         =   -1  'True
      ToolTipText     =   "Add File To Playlist"
      Top             =   1920
      Width           =   345
   End
   Begin VB.Image ImgAddFon 
      Height          =   300
      Left            =   7920
      Picture         =   "fmrMain.frx":29F1A
      Top             =   6240
      Width           =   345
   End
   Begin VB.Image ImgAddFoff 
      Height          =   300
      Left            =   7920
      Picture         =   "fmrMain.frx":2A4FC
      Top             =   5880
      Width           =   345
   End
   Begin VB.Image cmdExit 
      Height          =   300
      Left            =   3000
      Picture         =   "fmrMain.frx":2AADE
      Stretch         =   -1  'True
      ToolTipText     =   "Quit"
      Top             =   960
      Width           =   345
   End
   Begin VB.Image ImgExiton 
      Height          =   300
      Left            =   7920
      Picture         =   "fmrMain.frx":2B0C0
      Top             =   1920
      Width           =   360
   End
   Begin VB.Image ImgExitoff 
      Height          =   300
      Left            =   7920
      Picture         =   "fmrMain.frx":2B6A2
      Top             =   1560
      Width           =   360
   End
   Begin VB.Image ImgSfronton 
      Height          =   300
      Left            =   7920
      Picture         =   "fmrMain.frx":2BC84
      Top             =   3360
      Width           =   375
   End
   Begin VB.Image ImgSfrontoff 
      Height          =   300
      Left            =   7920
      Picture         =   "fmrMain.frx":2C2B6
      Top             =   3000
      Width           =   375
   End
   Begin VB.Image cmdRW 
      Height          =   300
      Left            =   120
      Picture         =   "fmrMain.frx":2C8E8
      Stretch         =   -1  'True
      ToolTipText     =   "Rewind"
      Top             =   1920
      Width           =   345
   End
   Begin VB.Image ImgSbackon 
      Height          =   300
      Left            =   7920
      Picture         =   "fmrMain.frx":2CF1A
      Top             =   4080
      Width           =   375
   End
   Begin VB.Image ImgSbackoff 
      Height          =   300
      Left            =   7920
      Picture         =   "fmrMain.frx":2D54C
      Top             =   3720
      Width           =   375
   End
   Begin VB.Image cmdNextSong 
      Height          =   300
      Left            =   3720
      Picture         =   "fmrMain.frx":2DB7E
      Stretch         =   -1  'True
      ToolTipText     =   "Next Track"
      Top             =   960
      Width           =   360
   End
   Begin VB.Image ImgNexton 
      Height          =   300
      Left            =   7920
      Picture         =   "fmrMain.frx":2E160
      Top             =   5520
      Width           =   345
   End
   Begin VB.Image ImgNextoff 
      Height          =   300
      Left            =   7920
      Picture         =   "fmrMain.frx":2E742
      Top             =   5160
      Width           =   345
   End
   Begin VB.Image cmdStop 
      Height          =   300
      Left            =   1200
      Picture         =   "fmrMain.frx":2ED24
      Stretch         =   -1  'True
      ToolTipText     =   "Stop"
      Top             =   960
      Width           =   345
   End
   Begin VB.Image ImgStopon 
      Height          =   300
      Left            =   9000
      Picture         =   "fmrMain.frx":2F306
      Top             =   480
      Width           =   345
   End
   Begin VB.Image ImgStopoff 
      Height          =   300
      Left            =   9000
      Picture         =   "fmrMain.frx":2F8E8
      Top             =   120
      Width           =   345
   End
   Begin VB.Image cmdPause 
      Height          =   300
      Left            =   840
      Picture         =   "fmrMain.frx":2FECA
      Stretch         =   -1  'True
      ToolTipText     =   "Pause"
      Top             =   960
      Width           =   345
   End
   Begin VB.Image ImgPauseon 
      Height          =   300
      Left            =   9360
      Picture         =   "fmrMain.frx":304AC
      Top             =   480
      Width           =   345
   End
   Begin VB.Image ImgPauseoff 
      Height          =   300
      Left            =   9360
      Picture         =   "fmrMain.frx":30A8E
      Top             =   120
      Width           =   345
   End
   Begin VB.Image cmdPlay 
      Height          =   300
      Left            =   480
      Picture         =   "fmrMain.frx":31070
      Stretch         =   -1  'True
      ToolTipText     =   "Play"
      Top             =   960
      Width           =   345
   End
   Begin VB.Image ImgPlayon 
      Height          =   300
      Left            =   8640
      Picture         =   "fmrMain.frx":31652
      Top             =   1200
      Width           =   345
   End
   Begin VB.Image ImgPlayoff 
      Height          =   300
      Left            =   8640
      Picture         =   "fmrMain.frx":31C34
      Top             =   840
      Width           =   345
   End
   Begin VB.Image ImgPrevon 
      Height          =   300
      Left            =   8640
      Picture         =   "fmrMain.frx":32216
      Top             =   480
      Width           =   345
   End
   Begin VB.Image ImgPrevoff 
      Height          =   300
      Left            =   8640
      Picture         =   "fmrMain.frx":327F8
      Top             =   120
      Width           =   345
   End
   Begin VB.Image cmdPrevSong 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   120
      Picture         =   "fmrMain.frx":32DDA
      Stretch         =   -1  'True
      ToolTipText     =   "Previous Track"
      Top             =   960
      Width           =   345
   End
   Begin VB.Label lblCenter 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Center"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   5040
      TabIndex        =   12
      Top             =   1560
      Width           =   630
   End
   Begin VB.Label LabelVolume 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "100%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   4320
      TabIndex        =   11
      Top             =   1560
      Width           =   630
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   375
      Left            =   9240
      TabIndex        =   10
      Top             =   840
      Visible         =   0   'False
      Width           =   375
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   0
      WindowlessVideo =   0   'False
   End
   Begin VB.Shape OutPut_Window 
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   1140
      Index           =   0
      Left            =   120
      Top             =   120
      Width           =   3975
   End
   Begin VB.Shape OutPut_Window 
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   4575
      Index           =   1
      Left            =   120
      Top             =   1920
      Width           =   3975
   End
   Begin VB.Image EqImage 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2055
      Left            =   4320
      Picture         =   "fmrMain.frx":333BC
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   2775
   End
   Begin VB.Menu mnuShortcuts 
      Caption         =   "Shortcuts"
      Visible         =   0   'False
      Begin VB.Menu mnuPlayer 
         Caption         =   "Player"
         Begin VB.Menu mnuPlay 
            Caption         =   "Play"
         End
         Begin VB.Menu mnuPause 
            Caption         =   "Pause"
         End
         Begin VB.Menu mnuStop 
            Caption         =   "Stop"
         End
         Begin VB.Menu mnuRewind 
            Caption         =   "Rewind"
         End
         Begin VB.Menu mnuForward 
            Caption         =   "Forward"
         End
         Begin VB.Menu mnuSpace1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuNextSong 
            Caption         =   "Next Song"
         End
         Begin VB.Menu mnuPreviousSong 
            Caption         =   "Previous Song"
         End
      End
      Begin VB.Menu mnuSpeakers 
         Caption         =   "Speakers"
         Begin VB.Menu mnuMuteSpeakers 
            Caption         =   "Mute Speakers"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuCenterVolume 
            Caption         =   "Center Volume"
         End
         Begin VB.Menu mnuCenterTreble 
            Caption         =   "Center Treble"
         End
         Begin VB.Menu mnuCenterBass 
            Caption         =   "Center Bass"
         End
      End
      Begin VB.Menu mnuSequence 
         Caption         =   "Sequence"
         Begin VB.Menu mnuSequenceNextSong 
            Caption         =   "Play Next Song"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuSequenceRnd 
            Caption         =   "Play Random Next Song"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuSequenceNone 
            Caption         =   "Stop after Playing"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuSettings 
         Caption         =   "Settings"
         Begin VB.Menu mnuOptions 
            Caption         =   "Options Form"
         End
         Begin VB.Menu mnuOnTop 
            Caption         =   "Always on Top"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuWinVol 
            Caption         =   "Show Win Volume"
         End
      End
      Begin VB.Menu mnuRestore 
         Caption         =   "Restore"
      End
      Begin VB.Menu mnuSpace2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuListOperations 
      Caption         =   "List Operations"
      Visible         =   0   'False
      Begin VB.Menu mnuSortAZ 
         Caption         =   "Sort List (A to Z)"
      End
      Begin VB.Menu mnuSortZA 
         Caption         =   "Sort List (Z to A)"
      End
      Begin VB.Menu mnuSpace4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cut Entry"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy Entry"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste Entry"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuCrop 
         Caption         =   "Crop Entry (*)"
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "Remove Entry (Del)"
      End
      Begin VB.Menu mnuSpace5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddDir 
         Caption         =   "Add Directory"
      End
      Begin VB.Menu mnuAddFile 
         Caption         =   "Add File (Ins)"
      End
      Begin VB.Menu mnuSpace6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWinVol2 
         Caption         =   "Show Win Volume"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-------------------------'
'    EQUALIZER PRESETS    '
'-------------------------'
Private Sub cmdAmpPreset_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Equalizer presets button update
    cmdAmpPreset.Picture = ImgEQpresetson.Picture
End Sub
Private Sub cmdAmpPreset_Click()
' Equalizer Presets
    ListPresets.Visible = True
    ListPresets.SetFocus
End Sub
Private Sub cmdAmpPreset_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Equalizer presets button update
    cmdAmpPreset.Picture = ImgEQpresetsoff.Picture
End Sub
Private Sub ListPresets_KeyDown(KeyCode As Integer, Shift As Integer)
' Handle list via arrows
    If KeyCode = vbKeyReturn Then ListPresets_DblClick
End Sub
Private Sub ListPresets_DblClick()
' Apply presets
SET_PRESET ListPresets.ListIndex
' Hide presets list
ListPresets.Visible = False
End Sub
'-------------------------'
'LOADING FORM/INITIALIZING'
'-------------------------'
Private Sub Form_Resize()
' Clear Caption
    Me.Caption = ""
    If (GREQ_ENABLED = True) And (FORM_TOOLS = True) Then
        frmGrEq.Top = frmMain.Top + 120
        frmGrEq.Left = frmMain.Left + 4320
        frmGrEq.Show 0, Me
    End If
End Sub
Private Sub Form_Load()
    mnuMuteSpeakers.Checked = False
    frmMain.Equ.Enabled = False
    For I = 0 To 4 Step 1
        frmMain.Preamp(I).Enabled = False
    Next I
    If vol.VolTrebleMax = 0 Then
        ' Tremble not functioning
        lblTreble.Caption = "N/A"
        lblTreble.ForeColor = RGB(255, 0, 0)
        sldMainTreb.Enabled = False
    Else
        sldMainTreb.min = vol.VolTrebleMin
        sldMainTreb.Max = vol.VolTrebleMax
        sldMainTreb.TickFrequency = (sldMainTreb.Max - sldMainTreb.min) \ 10
        sldMainTreb.LargeChange = sldMainTreb.TickFrequency
        If vol.VolBassMax = 0 Then
            ' Bass not functioning
            lblBass.Caption = "N/A"
            lblBass.ForeColor = RGB(255, 0, 0)
            sldMainBas.Enabled = False
        Else
            sldMainBas.min = vol.VolBassMin
            sldMainBas.Max = vol.VolBassMax
            sldMainBas.TickFrequency = (sldMainBas.Max - sldMainBas.min) \ 10
            sldMainBas.LargeChange = sldMainBas.TickFrequency
            ' Initialize Mixer only is bass/tremble working
            INITIALIZE_EQUALIZER
        End If
    End If
    INITIALIZE_LEVELS
End Sub
'-------------------------'
'  SHOW/HIDE TOOL AREA    '
'-------------------------'
Public Sub cmdTools_Click()
' Show tools
    If FORM_MODE = False Then
        If FORM_TOOLS = False Then
            For I = 1 To 15
                frmMain.Width = 4215 + I * 203
                WAIT (0.05)
            Next I
            ' Show equalizer
            GREQ_ENABLED = True
            frmGrEq.Top = frmMain.Top + 120
            frmGrEq.Left = frmMain.Left + 4320
            frmGrEq.Show 0, Me
            frmMain.Width = 7260
            cmdTools.Picture = ImgToolson.Picture
            FORM_TOOLS = True
        Else
            ' Hide equalizer
            GREQ_ENABLED = True
            frmGrEq.Hide
            For I = 1 To 15
                frmMain.Width = 7260 - I * 203
                WAIT (0.05)
            Next I
            frmMain.Width = 4215
            cmdTools.Picture = ImgToolsoff.Picture
            FORM_TOOLS = False
        End If
    End If
End Sub
'-------------------------'
'PLAY/INFORM VIA SONGLIST '
'-------------------------'
Private Sub List1_KeyUp(KeyCode As Integer, Shift As Integer)
' Update file number by pressing up arrow
    ListInfo.Text = List1.ListIndex + 1 & "/" & List1.ListCount & " File"
    List1.ToolTipText = List1.Text
End Sub
Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Show pop-up menu if right-clicked
    List1.ToolTipText = List1.Text
    If Button = vbRightButton Then
        PopupMenu Me.mnuListOperations
    End If
    Getmp3data (List1.Text)
End Sub
Public Sub List1_Click()
' Update file number by simple click
    ListInfo.Text = List1.ListIndex + 1 & "/" & List1.ListCount & " File"
    List1.ToolTipText = List1.Text
    Getmp3data (List1.Text)
End Sub
Private Sub List1_DblClick()
' Play file via playlist with doubleclick
    CmdPlay_Click
    List1.ToolTipText = List1.Text
    Getmp3data (List1.Text)
End Sub
Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
' Actions by the keyboard
    Select Case KeyCode
    Case Is = vbKeyDelete
        ' Remove entry with Del Key
        CmdRem_Click
    Case Is = vbKeyReturn
        ' Play entry with Enter Key
        CmdPlay_Click
    Case Is = vbKeyRight
        ' FF with right arrow
        cmdFF_Click
        List1.ListIndex = List1.ListIndex - 1
    Case Is = vbKeyLeft
        ' Rewind with left arrow
        cmdRW_Click
        List1.ListIndex = List1.ListIndex + 1
    Case Is = vbKeyInsert
        ' Rewind with left arrow
        cmdAddFile_Click
    Case Is = vbKeyMultiply
        ' Crop entry
        mnuCrop_Click
    End Select
    ListInfo.Text = List1.ListIndex + 1 & "/" & List1.ListCount & " File"
    List1.ToolTipText = List1.Text
    Getmp3data (List1.Text)
End Sub
'-------------------------'
'    MOVING THE FORM      '
'-------------------------'
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Realise the moving will by pressing a mouse key
    MOVE_FORM = True
    If GREQ_ENABLED = True Then
        frmGrEq.Top = frmMain.Top + 120
        frmGrEq.Left = frmMain.Left + 4320
    End If
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' System Tray Controls & move form if key pressed on it
Dim Result As Long
Dim msg As Long
    If GREQ_ENABLED = True Then
        frmGrEq.Top = frmMain.Top + 120
        frmGrEq.Left = frmMain.Left + 4320
    End If
    If IS_MINIMIZED = True Then
        If Me.ScaleMode = vbPixels Then
            msg = X
        Else
            msg = X / Screen.TwipsPerPixelX
        End If
        Select Case msg
        ' Button pressed while in systray mode
            Case WM_RBUTTONUP
                ' Left Button - Menu
                Result = SetForegroundWindow(Me.hwnd)
                Me.PopupMenu Me.mnuShortcuts
            Case WM_LBUTTONUP
                ' Right Button - Restore
                mnuRestore_Click
        End Select
    Else
        If MOVE_FORM = True Then
            ' Move form
            frmMain.Move frmMain.Left + (X - PREVIOUS_X), frmMain.Top + (Y - PREVIOUS_Y)
        Else
            PREVIOUS_X = X
            PREVIOUS_Y = Y
        End If
    End If
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Stop moving will
    MOVE_FORM = False
    If GREQ_ENABLED = True Then
        frmGrEq.Top = frmMain.Top + 120
        frmGrEq.Left = frmMain.Left + 4320
    End If
End Sub
'-------------------------'
'   PREVIOUS SONG BUTTON  '
'-------------------------'
Private Sub cmdPrevSong_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Previous song button update
    cmdPrevSong.Picture = ImgPrevon.Picture
End Sub
Private Sub cmdPrevSong_Click()
' Play previous song after the click on the button
   On Error Resume Next
    If PAUSED_PLAYER = True Then
    ' Check if player is paused to unpause it
        CmdPause_Click
    End If
    If List1.ListIndex <> 0 Then
        List1.ListIndex = List1.ListIndex - 1
        CmdPlay_Click
        Exit Sub
    Else
        List1.ListIndex = List1.ListCount - 1
        CmdPlay_Click
    End If
    List1.ToolTipText = List1.Text
End Sub
Private Sub cmdPrevSong_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Previous song button update
    cmdPrevSong.Picture = ImgPrevoff.Picture
End Sub
'-------------------------'
'     PLAY SONG BUTTON    '
'-------------------------'
Private Sub cmdPlay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Play song button update
    cmdPlay.Picture = ImgPlayon.Picture
End Sub
Public Sub CmdPlay_Click()
' Play track after the click on the button
    If PAUSED_PLAYER = True Then
    ' Check if player is paused to unpause it
        CmdPause_Click
    End If
    TxtName.Text = List1.Text
    On Error Resume Next
    MediaPlayer1.FileName = TxtName.Text
    If TxtName.Text <> "" Then
        MediaPlayer1.Play
        sldSearch.Max = MediaPlayer1.Duration
        SPLIT_MINUTES_SECONDS (MediaPlayer1.Duration)
        cmdPause.Enabled = True
        ImgAction.Picture = ImgPlayS.Picture
    End If
    FORM_TITLE = "Rasputin - " & List1.Text & " - " & ListInfo.Text & " (" & MINUTES_LEFT & ":" & SECONDS_LEFT & ") *** "
    SysTrayTitle = FORM_TITLE
    FIX_TITLE_BAR
End Sub
Private Sub cmdPlay_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Play song button update
    cmdPlay.Picture = ImgPlayoff.Picture
End Sub
'-------------------------'
'       PAUSE BUTTON      '
'-------------------------'
Private Sub CmdPause_Click()
' Pause song after the click on the button
    If List1.ListCount = 0 Then Exit Sub
    If TxtName.Text = "" Then Exit Sub
    If PAUSED_PLAYER = False Then
        MediaPlayer1.Pause
        PAUSED_PLAYER = True
        cmdPause.Picture = ImgPauseon.Picture
        FORM_TITLE = "Rasputin - " & TxtName.Text & " - " & ListInfo.Text & "(" & MINUTES_LEFT & ":" & SECONDS_LEFT & ") (Paused) *** "
        SysTrayTitle = FORM_TITLE
        ImgAction.Picture = ImgPauseS.Picture
    Else
        MediaPlayer1.Play
        PAUSED_PLAYER = False
        cmdPause.Picture = ImgPauseoff.Picture
        FORM_TITLE = "Rasputin - " & TxtName.Text & " - " & ListInfo.Text & "(" & MINUTES_LEFT & ":" & SECONDS_LEFT & ") *** "
        SysTrayTitle = FORM_TITLE
    End If
End Sub
'-------------------------'
'       STOP  BUTTON      '
'-------------------------'
Private Sub CmdStop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Stop song button update
    cmdStop.Picture = ImgStopon.Picture
End Sub
Private Sub CmdStop_Click()
' Stop playing after the click on the button
    If PAUSED_PLAYER = True Then
    ' Check if player is paused to unpause it
        CmdPause_Click
    End If
    MediaPlayer1.Stop
    sldSearch.Value = 0
    FORM_TITLE = "Rasputin - " & TxtName.Text & " - " & ListInfo.Text & "(" & MINUTES_LEFT & ":" & SECONDS_LEFT & ") (Stopped) *** "
    SysTrayTitle = FORM_TITLE
    cmdPause.Enabled = False
    ImgAction.Picture = ImgStopS.Picture
End Sub
Private Sub CmdStop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Stop song button update
    cmdStop.Picture = ImgStopoff.Picture
End Sub
'-------------------------'
'     NEXT SONG BUTTON    '
'-------------------------'
Private Sub cmdNextSong_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Next song button update
    cmdNextSong.Picture = ImgNexton.Picture
End Sub
Private Sub cmdNextSong_Click()
' Play next song after the click on the button
    If PAUSED_PLAYER = True Then
    ' Check if player is paused to unpause it
        CmdPause_Click
    End If
    On Error Resume Next
    If List1.ListIndex < List1.ListCount - 1 Then
        List1.ListIndex = List1.ListIndex + 1
        CmdPlay_Click
        Exit Sub
    Else
        List1.ListIndex = 0
        CmdPlay_Click
    End If
    List1.ToolTipText = List1.Text
End Sub
Private Sub cmdNextSong_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Next song button update
    cmdNextSong.Picture = ImgNextoff.Picture
End Sub
'-------------------------'
'          REWIND         '
'-------------------------'
Private Sub cmdRW_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Rewind button update
    cmdRW.Picture = ImgSbackon.Picture
End Sub
Private Sub cmdRW_Click()
' Rewind after the click on the button
    sldSearch.Value = sldSearch.Value - RW
    sldSearch_Scroll
End Sub
Private Sub cmdRW_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Rewind button update
    cmdRW.Picture = ImgSbackoff.Picture
End Sub
'-------------------------'
'      FAST FORWARD       '
'-------------------------'
Private Sub cmdFF_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Fast forward button update
    cmdFF.Picture = ImgSfronton.Picture
End Sub
Private Sub cmdFF_Click()
' Fast forward after the click on the button
    sldSearch.Value = sldSearch.Value + FF
    sldSearch_Scroll
End Sub
Private Sub cmdFF_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Fast forward button update
    cmdFF.Picture = ImgSfrontoff.Picture
End Sub
'-------------------------'
'     QUIT Rasputin       '
'-------------------------'
Private Sub cmdExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Quit button update
    cmdExit.Picture = ImgExiton.Picture
End Sub
Private Sub cmdExit_Click()
' Quit and save current playlist to auto-load it later
    If AUTOLOAD_LIST = True Then
        AUTOSAVE_PLAYLIST
    End If
    SaveSetting App.Title, "Equalize", "T_LEVEL", sldMainTreb.Value
    SaveSetting App.Title, "Equalize", "B_LEVEL", sldMainBas.Value
    End
    Unload Me
End Sub
Private Sub cmdExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Quit button update
    cmdExit.Picture = ImgExitoff.Picture
End Sub
Private Sub Form_Terminate()
' Terminate form and reclaim greq memory
    If DevHandle <> 0 Then
        Call DoStop
    End If
    End
End Sub
Private Sub Form_Unload(Cancel As Integer)
' Unload Systray Options and reclaim greq memory
    Shell_NotifyIcon NIM_DELETE, SYSTRAYMODE
    If DevHandle <> 0 Then
        Call DoStop
    End If
    End
End Sub
'-------------------------'
'  ADD FILE TO PLAYLIST   '
'-------------------------'
Private Sub cmdAddFile_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Add file to playlist button update
    cmdAddFile.Picture = ImgAddFon.Picture
End Sub
Private Sub cmdAddFile_Click()
' Add file to playlist via common dialogues
    CommonDialog1.DialogTitle = "Load Sound File"
    CommonDialog1.MaxFileSize = 16384
    CommonDialog1.FileName = ""
    CommonDialog1.Filter = "MP3 Files (*.mp3)|*.MP3|Midi Files (*.mid)|*.mid|Wave Files (*.wav)|*.wav"
    CommonDialog1.ShowOpen
    If CommonDialog1.FileTitle <> "" Then
        List1.AddItem CommonDialog1.FileName
        TxtName.Text = CommonDialog1.FileName
        List1.ListIndex = List1.ListCount - 1
    End If
    List1_Click
    FIX_TITLE_BAR
    Exit Sub
End Sub
Private Sub cmdAddFile_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Add file to playlist button update
    cmdAddFile.Picture = ImgAddFoff.Picture
End Sub
'-------------------------'
'ADD DIRECTORY TO PLAYLIST'
'-------------------------'
Private Sub cmdAdddir_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Add directory button update
    cmdAdddir.Picture = ImgAddDon.Picture
End Sub
Private Sub cmdAdddir_Click()
' Prepare form for the new controls
    PREPARE_ADD_DIR
End Sub
Private Sub cmdAdddir_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Add directory button update
    cmdAdddir.Picture = ImgAddDoff.Picture
End Sub
Private Sub Drive1_Change()
' Change drive / dir path
    On Error Resume Next
    Dir1.Path = Drive1.Drive
End Sub
Private Sub Dir1_Change()
' Changer dir / file path
    File1.Path = Dir1.Path
End Sub
'-------------------------'
' CONFIRM DIR. ADDITTION  '
'-------------------------'
Private Sub cmdOK_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Add directory button update
    cmdOK.Picture = ImgConfirmon.Picture
End Sub
Private Sub cmdOK_Click()
' Add directory by pressing the OK button
    If File1.ListCount <> 0 Then
        For I = 1 To File1.ListCount
            File1.ListIndex = I - 1
            If Len(Dir1.Path) > 3 Then
                List1.AddItem Dir1.Path & "\" & File1.FileName
            Else
                List1.AddItem Dir1.Path & File1.FileName
            End If
        Next I
    End If
    RETURN_ADD_DIR
    List1_Click
    List1.ListIndex = List1.ListCount - 1
End Sub
Private Sub cmdOK_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Add directory button update
    cmdOK.Picture = ImgConfirmoff.Picture
End Sub
'--------------------------'
'CANCEL DIRECTORY ADDITTION'
'--------------------------'
Private Sub cmdCancel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Cancel add dir button update
    cmdCancel.Picture = ImgCancelon.Picture
End Sub
Private Sub cmdCancel_Click()
' Cancel adding directory and reform the form
    RETURN_ADD_DIR
End Sub
Private Sub cmdCancel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Cancel add dir button update
    cmdCancel.Picture = ImgCanceloff.Picture
End Sub
'-------------------------'
'REMOVE ONE PLAYLIST ENTRY'
'-------------------------'
Private Sub cmdRem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Remove 1 List Entry button update
    cmdRem.Picture = ImgRem1on.Picture
End Sub
Private Sub CmdRem_Click()
' Remove 1 List Entry after choosing the track from the playlist
    If List1.ListIndex <> -1 Then
        If List1.Text = TxtName.Text Then TxtName.Text = ""
        List1.RemoveItem List1.ListIndex
    End If
    List1_Click
End Sub
Private Sub cmdRem_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Remove 1 List Entry button update
    cmdRem.Picture = ImgRem1off.Picture
End Sub
'-------------------------'
'REMOVE ALL PLAYLIST SONGS'
'-------------------------'
Private Sub cmdClear_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Remove all entries button update
    cmdClear.Picture = ImgRemAon.Picture
End Sub
Private Sub CmdClear_Click()
' Remove all entries from the playlist
    List1.Clear
    TxtName.Text = ""
    ListInfo.Text = "0/0 File"
End Sub
Private Sub cmdClear_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Remove all entries button update
    cmdClear.Picture = ImgRemAoff.Picture
End Sub
'-------------------------'
'SAVE PLAYLIST  RPL FORMAT'
'-------------------------'
Private Sub CmdSaveList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Save playlist button update
    CmdSaveList.Picture = ImgSaveLon.Picture
End Sub
Private Sub CmdSaveList_Click()
' Save playlist to a text file with rpl extension
    Dim ListName As String
    CommonDialog3.DialogTitle = "Save Rasputin Playlist"
    CommonDialog3.MaxFileSize = 16384
    CommonDialog3.Filter = "Rasputin Playlists (*.rpl)|*.rpl"
    CommonDialog3.FileName = ""
    CommonDialog3.InitDir = App.Path
    CommonDialog3.DefaultExt = ".rpl"
    CommonDialog3.ShowSave
    If CommonDialog3.FileName = "" Then Exit Sub
    ListName = CommonDialog3.FileName
    If Right(ListName, 4) <> ".rpl" Then ListName = ListName + ".rpl"
    On Error GoTo Problem
    Open (ListName) For Output As #1
    Print #1, "Rasputin Playlist : " & ListName
    Dim I%
    For I = 0 To List1.ListCount - 1
    Print #1, List1.List(I)
    Next
    Close #1
    Exit Sub
Problem:
RESPONCE = MsgBox("An error appeared while trying to save" & vbNewLine & "the list and the file was not created.", vbOKOnly & vbInformation, "Error while saving playlist")
End Sub
Private Sub CmdSaveList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Save playlist button update
    CmdSaveList.Picture = ImgSaveLoff.Picture
End Sub
'-------------------------'
'LOAD PLAYLIST  RPL FORMAT'
'-------------------------'
Private Sub CmdLoadList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Load Playlist button update
    CmdLoadList.Picture = ImgLoadLon.Picture
End Sub
Private Sub CmdLoadList_Click()
' Load Playlist (text file with rpl extension)
    Dim ListName As String
    CommonDialog2.DialogTitle = "Load Rasputin Playlist"
    CommonDialog2.MaxFileSize = 16384
    CommonDialog2.FileName = ""
    CommonDialog2.Filter = "Rasputin Playlists (*.rpl)|*.rpl"
    CommonDialog2.ShowOpen
    If CommonDialog2.FileName = "" Then Exit Sub
    ListName = CommonDialog2.FileName
    On Error GoTo Problem
    Open ListName For Input As #1
    Input #1, SONG$
    If Left(SONG$, 20) <> "Rasputin Playlist : " Then GoTo Problem
    Do Until EOF(1)
        Input #1, SONG$
        List1.AddItem SONG$
    Loop
    Close 1
    List1.ListIndex = List1.ListCount - 1
    List1_Click
    Exit Sub
Problem:
    RESPONCE = MsgBox("The list appears to be corrupted or some files" & vbNewLine & "are missing from their original location.", vbOKOnly & vbInformation, "Error while loading playlist")
End Sub
Private Sub CmdLoadList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Load Playlist button update
    CmdLoadList.Picture = ImgLoadLoff.Picture
End Sub
'-------------------------'
'  MINIMIZE THE PROGRAM   '
'-------------------------'
Private Sub cmdMinimize_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Minimize button update
    cmdMinimize.Picture = ImgMinon.Picture
End Sub
Private Sub cmdMinimize_Click()
' Minimize the form
    frmGrEq.Hide
    IS_MINIMIZED = True
    If MINIMIZE_MODE = False Then
        Me.WindowState = vbMinimized
        Me.Caption = FORM_TITLE
    Else
        With SYSTRAYMODE
            .cbSize = Len(SYSTRAYMODE)
            .hwnd = Me.hwnd
            .uId = vbNull
            .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
            .uCallBackMessage = WM_MOUSEMOVE
            .hIcon = Me.Icon
            .szTip = SysTrayTitle & vbNullChar
        End With
        Shell_NotifyIcon NIM_ADD, SYSTRAYMODE
        frmMain.Hide
    End If
End Sub
Private Sub cmdMinimize_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Minimize button update
    cmdMinimize.Picture = ImgMinoff.Picture
End Sub
'-------------------------'
'NEXT SONG - NEXT/RND/NONE'
'-------------------------'
Private Sub cmdNextList_Click()
' Play Next song after finishing playing current song
    NEXT_SONG = 1
    SEQUENCE
End Sub
Private Sub cmdRndList_Click()
' Play random song after after finishing playing current song
    NEXT_SONG = 2
    SEQUENCE
End Sub
Private Sub cmdNoneList_Click()
' Dont play any song after after finishing playing current song
    NEXT_SONG = 3
    SEQUENCE
End Sub
'-------------------------'
'  KEEP MAIN FORM ON TOP  '
'-------------------------'
Public Sub CmdOnTop_Click()
' Always on top button
    If FORM_ON_TOP = False Then
        AlwaysOnTop frmMain, True
        cmdOnTop.Picture = ImgTopon.Picture
        FORM_ON_TOP = True
        mnuOnTop.Checked = True
    Else
        AlwaysOnTop frmMain, False
        cmdOnTop.Picture = ImgTopoff.Picture
        FORM_ON_TOP = False
        mnuOnTop.Checked = False
    End If
End Sub
Public Sub AlwaysOnTop(frmMain As Form, SetOnTop As Boolean)
' Set the form on top
    If SetOnTop Then
        MAINFORM_FLAG = HWND_TOPMOST
    Else
        MAINFORM_FLAG = HWND_NOTOPMOST
    End If
    SetWindowPos frmMain.hwnd, MAINFORM_FLAG, frmMain.Left / Screen.TwipsPerPixelX, _
    frmMain.Top / Screen.TwipsPerPixelY, frmMain.Width / Screen.TwipsPerPixelX, _
    frmMain.Height / Screen.TwipsPerPixelY, SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub
'-------------------------'
' MUTE/MANIPULATE SPEAKER '
'-------------------------'
Private Sub CmdMute_Click()
' Mute Speaker
    If MediaPlayer1.Mute = False Then
        MediaPlayer1.Mute = True
        volume = sldVol.Value
        sldVol.Value = 0
        LabelVolume.Caption = "0%"
        LabelVolume.ForeColor = RGB(155 + sldVol.Value / 10, 100, 100)
        cmdMute.Picture = ImgMuteon.Picture
        mnuMuteSpeakers.Checked = True
    Else
        MediaPlayer1.Mute = False
        sldVol.Value = volume
        sldVol_Scroll
        cmdMute.Picture = ImgMuteoff.Picture
        mnuMuteSpeakers.Checked = False
    End If
End Sub
Private Sub sldVol_Scroll()
' Change volume
    Dim MinVal As Integer, SldMin As Integer
    LabelVolume.ForeColor = RGB(155 + sldVol.Value / 10, 100, 100)
    TempVar = sldVol.Value - 2500
    MediaPlayer1.volume = TempVar
    On Error GoTo Problem
    SldMin = sldVol.min
    MinVal = sldVol.Value
    LabelVolume.Caption = MinVal \ 25 & " %"
Problem:
    Exit Sub
End Sub
Private Sub INITIALIZE_LEVELS()
' Adjust Bass / Treble to previous state in last use
    On Error Resume Next
    sldMainTreb.Value = Val(GetSetting(App.Title, "Equalize", "T_LEVEL", vol.VolTrebleMax))
        sldMainTreb_Scroll
    sldMainBas.Value = Val(GetSetting(App.Title, "Equalize", "B_LEVEL", vol.VolBassMax))
        sldMainBas_Scroll
End Sub
'-------------------------'
'CENTRALIZE/REMORPH ASPECT'
'-------------------------'
Private Sub cmdCenter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Center volume button update
    cmdCenter.Picture = ImgCenteron.Picture
End Sub
Private Sub cmdCenter_Click()
' Center volume
    sldCenter.Value = 0
    sldCenter_Scroll
End Sub
Private Sub cmdCenter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Center volume button update
    cmdCenter.Picture = ImgCenteroff.Picture
End Sub
Private Sub sldCenter_Scroll()
' Change the centralization of the sound
    On Error GoTo Problem
    Select Case sldCenter.Value
    Case Is = 0
        lblCenter.Caption = "Center"
        cmdCenter.Picture = ImgCenteron.Picture
    Case Is < 0
        If sldCenter.Value > -5000 Then
            lblCenter.Caption = "L " & Int(-sldCenter.Value / 50) & "%"
        Else
            lblCenter.Caption = "Left"
        End If
        cmdCenter.Picture = ImgCenteroff.Picture
    Case Is > 0
        If sldCenter.Value < 5000 Then
            lblCenter.Caption = "R " & Int(sldCenter.Value / 50) & "%"
        Else
            lblCenter.Caption = "Right"
        End If
        cmdCenter.Picture = ImgCenteroff.Picture
    End Select
    MediaPlayer1.Balance = sldCenter.Value
    Exit Sub
Problem:
    Exit Sub
End Sub
'-------------------------'
'   MEDIA PLAYER TIMING   '
'-------------------------'
Private Sub TimerPlayer_Timer()
' Media Player Timing
    Dim min As Integer
    Dim sec As Integer
    On Error Resume Next
    ' Autoplay_check
    If AUTOPLAY = True Then
        CmdPlay_Click
        AUTOPLAY = False
    End If
    sldSearch.Value = MediaPlayer1.CurrentPosition
    SecTemp = MediaPlayer1.CurrentPosition
    min = SecTemp \ 60
    sec = SecTemp - (min * 60)
    If sec = "-1" Then sec = "0"
    ' Show minutes
    If sec < 10 Then
        TxtTime.Text = min & ":0" & sec & "/" & MINUTES_LEFT & ":" & SECONDS_LEFT
    Else
        TxtTime.Text = min & ":" & sec & "/" & MINUTES_LEFT & ":" & SECONDS_LEFT
    End If
    ' Show bytes
    If frmMain.Width = 7260 Then
        If MediaPlayer1.CurrentPosition <> -1 Then
            lblBytes.Caption = MediaPlayer1.CurrentPosition & " / " & MediaPlayer1.Duration
        Else
            lblBytes.Caption = ""
        End If
    End If
End Sub
Private Sub MediaPlayer1_EndOfStream(ByVal Result As Long)
' Handling media player when finishing a track
    On Error Resume Next
    If NEXT_SONG = 2 Then
        ' Play next song
        Randomize Timer
        TempVar = Int((List1.ListCount * Rnd))
        List1.ListIndex = TempVar
        CmdPlay_Click
    Else
        If NEXT_SONG = 1 Then
        ' Play random song
            If List1.ListIndex < List1.ListCount - 1 Then
                List1.ListIndex = List1.ListIndex + 1
                CmdPlay_Click
                Exit Sub
            Else
            ' Don't playing anything
                MediaPlayer1.Stop
            End If
        End If
    End If
    If NEXT_SONG = 3 Then
        MediaPlayer1.Stop
    End If
End Sub
'-------------------------'
'SLIDER USAGE IN SEARCHING'
'-------------------------'
Private Sub sldSearch_Scroll()
' Search the track via the scroll bar
    MediaPlayer1.CurrentPosition = sldSearch.Value
End Sub
'-------------------------'
'  HANDLE THE BASS LEVEL  '
'-------------------------'
Private Sub cmdBassC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Center bass button update
    cmdBassC.Picture = ImgCenteron.Picture
End Sub
Private Sub cmdBassC_Click()
' Centralize the Bass level
    sldMainBas.Value = Int(sldMainBas.Max / 2)
    sldMainBas_Scroll
End Sub
Private Sub sldMainBas_Scroll()
' Change Bass Level via the scroll bar
    vol.VolumeLevelBass = sldMainBas.Value
    lblBass.Caption = "B " & sldMainBas.Value
    lblBass.ForeColor = RGB(Int(sldMainBas.Value / 257), 100, 100)
End Sub
Private Sub cmdBassC_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Center bass button update
    cmdBassC.Picture = ImgCenteroff.Picture
End Sub
'-------------------------'
'HANDLE THE TREMBLE LEVEL '
'-------------------------'
Private Sub cmdTreC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Center Treble button update
    cmdTreC.Picture = ImgCenter2on.Picture
End Sub
Private Sub cmdTreC_Click()
' Centralize the Treble level
    sldMainTreb.Value = Int(sldMainTreb.Max / 2)
    sldMainTreb_Scroll
End Sub
Private Sub sldMainTreb_Scroll()
' Change Treble Level via the scrollbar
    vol.VolumeLevelTreble = sldMainTreb.Value
    lblTreble.Caption = "T " & sldMainTreb.Value
    lblTreble.ForeColor = RGB(Int(sldMainTreb.Value / 257), 100, 100)
End Sub
Private Sub cmdTreC_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Center Treble button update
    cmdTreC.Picture = ImgCenter2off.Picture
End Sub
'-------------------------'
' DIRECTORY ONE LEVEL UP  '
'-------------------------'
Private Sub cmdLup_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Directory Level Up button update
    cmdLup.Picture = ImgLupon.Picture
End Sub
Private Sub cmdLup_Click()
' Directory Level Up in the Dir List Box
    On Error Resume Next
    For I = 1 To Len(Dir1.Path)
        If Mid(Dir1.Path, I, 1) = "\" Then TempVar = I
    Next
    Dir1.Path = Left(Dir1.Path, TempVar)
    Dir1_Change
End Sub
Private Sub cmdLup_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Directory Level Up button update
    cmdLup.Picture = ImgLupoff.Picture
End Sub
'-------------------------'
'GO TO MY DOCUMENTS FOLDER'
'-------------------------'
Private Sub cmdMyDocuments_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Go to My Documents button update
    cmdMyDocuments.Picture = ImgMyon.Picture
End Sub
Private Sub cmdMyDocuments_Click()
' Go to My documents
    On Error GoTo GREEK_SYSTEM
    Dir1.Path = Left(WINDOWS_PATH, 3) & "My Documents"
    Drive1.Drive = Left(WINDOWS_PATH, 1)
    Exit Sub
GREEK_SYSTEM:
    ' Support the Greek Windows Version
    Dir1.Path = Left(WINDOWS_PATH, 3) & "Ôá ¸ããñáöÜ ìïõ"
End Sub
Private Sub cmdMyDocuments_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Go to My Documents button update
    cmdMyDocuments.Picture = ImgMyoff.Picture
End Sub
'-------------------------'
'   GO TO DESKTOP AREA    '
'-------------------------'
Private Sub cmdDesktop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Go to Desktop button update
    cmdDesktop.Picture = ImgDeskon.Picture
End Sub
Private Sub cmdDesktop_Click()
    On Error Resume Next
    Dir1.Path = WINDOWS_PATH & "\Desktop"
    Drive1.Drive = Left(WINDOWS_PATH, 1)
End Sub
Private Sub cmdDesktop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Go to Desktop button update
    cmdDesktop.Picture = ImgDeskoff.Picture
End Sub
'--------------------------'
'SCROLL TITLE/UPDATE SCREEN'
'--------------------------'
Private Sub TimerTitle_Timer()
' Scroll the title is it is prefered by the user
    If TITLE_SCROLL Then
        FORM_TITLE = Right(FORM_TITLE, Len(FORM_TITLE) - 1) & Left(FORM_TITLE, 1)
        TxtName.Text = FORM_TITLE
'        frmMain.Caption = FORM_TITLE
    Else
        TxtName.Text = FORM_TITLE
'        frmMain.Caption = FORM_TITLE
    End If
    ' Check condition
    Select Case MediaPlayer1.PlayState
    Case mpClosed
        ImgAction.Picture = ImgNull.Picture
    Case mpPaused
        ImgAction.Picture = ImgPauseS.Picture
    Case mpPlaying
        ImgAction.Picture = ImgPlayS.Picture
    Case mpStopped
        ImgAction.Picture = ImgStopS.Picture
    End Select
End Sub
'--------------------------------------------------'
'FORM MODE (COMPACT WITHOUT LIST OR FULL WITH LIST)'
'--------------------------------------------------'
Private Sub Form_DblClick()
' Double click to resize
    cmdCompact_Click
End Sub
Public Sub cmdCompact_Click()
' Compact mode (without list) / List mode (with list)
If FORM_MODE = False Then
    ' Hide equalizer
    GREQ_ENABLED = False
    frmGrEq.Hide
    ' Hide tools also
    If FORM_TOOLS = True Then
        For I = 1 To 15
            frmMain.Width = 7260 - I * 203
            WAIT (0.05)
        Next I
        frmMain.Width = 4215
        cmdTools.Picture = ImgToolsoff.Picture
    End If
    ' Pull down effect
    For I = 1 To 15
        frmMain.Height = 6630 - I * 350
        WAIT (0.05)
    Next I
    frmMain.Height = 1380
    cmdCompact.Picture = ImgMaxon.Picture
    FORM_MODE = True
    cmdCompact.ToolTipText = "Full mode"
Else
    ' Pull down effect
    For I = 1 To 15
        frmMain.Height = 1380 + I * 350
        WAIT (0.05)
    Next I
    frmMain.Height = 6630
    cmdCompact.Picture = ImgMaxoff.Picture
    FORM_MODE = False
    cmdCompact.ToolTipText = "Compact mode"
    ' Show tools if they were available
    If FORM_TOOLS = True Then
            For I = 1 To 15
                frmMain.Width = 4215 + I * 203
                WAIT (0.05)
            Next I
        ' Show equalizer
        GREQ_ENABLED = True
        frmGrEq.Top = frmMain.Top + 120
        frmGrEq.Left = frmMain.Left + 4320
        frmGrEq.Show 0, Me
        frmMain.Width = 7260
        cmdTools.Picture = ImgToolson.Picture
    End If
End If
End Sub
'-------------------------'
'  LOADING OPTIONS FORM   '
'-------------------------'
Private Sub cmdControls_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Controls button update
    cmdControls.Picture = ImgControlson.Picture
End Sub
Private Sub cmdControls_Click()
    ' Load options form
    frmOptions.Show
End Sub
Private Sub cmdControls_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Controls button update
    cmdControls.Picture = ImgControlsoff.Picture
End Sub
'-------------------------'
'ADJUST EQUALIZER MAIN BAR'
'-------------------------'
Private Sub Equ_Change()
' Adjust bass and treble
On Error Resume Next
    sldMainTreb.Value = sldMainTreb.Value + (Equ.Value - EquTemp)
    sldMainTreb_Scroll
    sldMainBas.Value = sldMainBas.Value + (Equ.Value - EquTemp)
    sldMainBas_Scroll
    EquTemp = Equ.Value
End Sub
'-------------------------'
' ADJUST EQUALIZER 5 BARS '
'-------------------------'
Private Sub Preamp_Change(Index As Integer)
' Adjust bass and tremble for 5 bars
On Error Resume Next
    If (Preamp(Index).Value - Equ5temp(Index)) <> 0 Then
        sldMainTreb.Value = sldMainTreb.Value + ((Preamp(Index).Value - Equ5temp(Index)) * (Index + 1) * 50)
        sldMainTreb_Scroll
        sldMainBas.Value = sldMainBas.Value + ((Preamp(Index).Value - Equ5temp(Index)) * (5 - Index) * 50)
        sldMainBas_Scroll
        Equ5temp(Index) = Preamp(Index).Value
    End If
End Sub
'--------------------------'
'   SAVE EQUALIZER STATE   '
'--------------------------'
Private Sub cmdAmpSave_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Save Amplifier button update
    cmdAmpSave.Picture = ImgEqSaveon.Picture
End Sub
Private Sub cmdAmpSave_Click()
' Save Equalizer Presets
    ListPresets.Visible = False
    Dim ListName As String
    CommonDialog3.DialogTitle = "Save Rasputin Equalizer State"
    CommonDialog3.MaxFileSize = 16384
    CommonDialog3.Filter = "Rasputin Equalizer (*.req)|*.req"
    CommonDialog3.FileName = ""
    CommonDialog3.InitDir = App.Path
    CommonDialog3.DefaultExt = ".req"
    CommonDialog3.ShowSave
    If CommonDialog3.FileName = "" Then Exit Sub
    ListName = CommonDialog3.FileName
    If Right(ListName, 4) <> ".req" Then ListName = ListName + ".req"
    On Error GoTo Problem
    Open (ListName) For Output As #1
    Print #1, "Rasputin Equalizer : " & ListName
    For I = 0 To 4 Step 1
        Print #1, Str(Preamp(I).Value)
    Next I
    Close #1
    Exit Sub
Problem:
RESPONCE = MsgBox("An error appeared while trying to save" & vbNewLine & "the equalizer and the file was not created.", vbOKOnly & vbInformation, "Error while saving equalizer")
End Sub
Private Sub cmdAmpSave_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Save Amplifier button update
    cmdAmpSave.Picture = ImgEqSaveoff.Picture
End Sub
'--------------------------'
'    DEFAULT EQALIZER      '
'--------------------------'
Private Sub cmdAmpDefault_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Default EQ button Update
    cmdAmpDefault.Picture = ImgEQdefon.Picture
End Sub
Private Sub cmdAmpDefault_Click()
' Default values to Preamp
    ListPresets.Visible = False
    For I = 0 To 4 Step 1
        Preamp(I).Value = 0
    Next I
    Equ.Value = 0
End Sub
Private Sub cmdAmpDefault_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Default EQ button Update
    cmdAmpDefault.Picture = ImgEQdefoff.Picture
End Sub
'--------------------------'
'   LOAD EQUALIZER STATE   '
'--------------------------'
Private Sub cmdAmpLoad_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Load Equalizer button update
    cmdAmpLoad.Picture = ImgEQloadon.Picture
End Sub
Private Sub cmdAmpLoad_Click()
' Load Equalizer Presets
    ListPresets.Visible = False
    Dim ListName As String
    Dim Eqload As String
    CommonDialog2.DialogTitle = "Load Rasputin Equalizer State"
    CommonDialog2.MaxFileSize = 16384
    CommonDialog2.FileName = ""
    CommonDialog2.Filter = "Rasputin Equalizer (*.req)|*.req"
    CommonDialog2.ShowOpen
    If CommonDialog2.FileName = "" Then Exit Sub
    ListName = CommonDialog2.FileName
    On Error GoTo Problem
    Open ListName For Input As #1
    Input #1, SONG$
    If Left(SONG$, 21) <> "Rasputin Equalizer : " Then GoTo Problem
    For I = 0 To 4 Step 1
        Line Input #1, Eqload
        If Left(Eqload, 1) = " " Then Eqload = Val(Mid(Eqload, 2, Len(Eqload) - 1))
        Preamp(I).Value = Val(Eqload)
    Next I
    Close 1
    Exit Sub
Problem:
    RESPONCE = MsgBox("The file appears to be corrupted or some data" & vbNewLine & "are missing from their original location.", vbOKOnly & vbInformation, "Error while loading equalizer")
End Sub
Private Sub cmdAmpLoad_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Load Equalizer button update
    cmdAmpLoad.Picture = ImgEQloadoff.Picture
End Sub
'------------------------------'
'SHORTCUTS FOR SYSTEM TRAY MODE'
'------------------------------'
Private Sub mnuExit_Click()
' Exit Program
    On Error Resume Next
    cmdExit_Click
End Sub
Private Sub mnuOptions_Click()
' Show Options form
    On Error Resume Next
    cmdControls_Click
End Sub
Private Sub mnuCenterBass_Click()
' Center bass
    On Error Resume Next
    cmdBassC_Click
End Sub
Private Sub mnuCenterTreble_Click()
' Center treble
    On Error Resume Next
    cmdTreC_Click
End Sub
Private Sub mnuCenterVolume_Click()
' Center volume
    On Error Resume Next
    cmdCenter_Click
End Sub
Private Sub mnuMuteSpeakers_Click()
' Mute
    On Error Resume Next
    CmdMute_Click
End Sub
Private Sub mnuPlay_Click()
' Play
    On Error Resume Next
    CmdPlay_Click
    If mnuPause.Caption = "Resume" Then
        mnuPause.Caption = "Pause"
        mnuPlay.Enabled = True
    End If
End Sub
Private Sub mnuStop_Click()
' Stop playing
    On Error Resume Next
    CmdStop_Click
    If mnuPause.Caption = "Resume" Then
        mnuPause.Caption = "Pause"
        mnuPlay.Enabled = True
    End If
End Sub
Private Sub mnuPause_Click()
' Pause
    On Error Resume Next
    CmdPause_Click
    If mnuPause.Caption = "Pause" Then
        mnuPause.Caption = "Resume"
        mnuPlay.Enabled = False
    Else
        mnuPause.Caption = "Pause"
        mnuPlay.Enabled = True
    End If
End Sub
Private Sub mnuForward_Click()
' Fast forward
    On Error Resume Next
    cmdFF_Click
End Sub
Private Sub mnuRewind_Click()
' Rewind
    On Error Resume Next
    cmdRW_Click
End Sub
Private Sub mnuNextSong_Click()
' Play Next Song
    On Error Resume Next
    cmdNextSong_Click
End Sub
Private Sub mnuPreviousSong_Click()
' Play Previous Song
    On Error Resume Next
    cmdPrevSong_Click
End Sub
Private Sub mnuRestore_Click()
' Restore Program
    On Error Resume Next
    IS_MINIMIZED = False
    Shell_NotifyIcon NIM_DELETE, SYSTRAYMODE
    Me.Show
End Sub
Private Sub mnuSequenceNextSong_Click()
' Play Next Song
    mnuSequenceNextSong.Checked = True
    mnuSequenceRnd.Checked = False
    mnuSequenceNone.Checked = False
    cmdNextList_Click
End Sub
Private Sub mnuSequenceNone_Click()
' Don't play a song next
    mnuSequenceNextSong.Checked = False
    mnuSequenceRnd.Checked = False
    mnuSequenceNone.Checked = True
    cmdNoneList_Click
End Sub
Private Sub mnuSequenceRnd_Click()
' Play random song next
    mnuSequenceNextSong.Checked = False
    mnuSequenceRnd.Checked = True
    mnuSequenceNone.Checked = False
    cmdRndList_Click
End Sub
Private Sub mnuOnTop_Click()
' Always on top
    CmdOnTop_Click
End Sub
'-----------------------------------'
'SHORTCUTS FOR PLAYLIST POP-UP MENOU'
'-----------------------------------'
Private Sub mnuRemove_Click()
' Remove Entry
    On Error Resume Next
    CmdRem_Click
End Sub
Private Sub mnuAddFile_Click()
' Add file
    On Error Resume Next
    cmdAddFile_Click
End Sub
Private Sub mnuAddDir_Click()
' Add directory
    On Error Resume Next
    cmdAdddir_Click
End Sub
Private Sub mnuCrop_Click()
' Crop entry
    On Error Resume Next
    List1.AddItem List1.Text, List1.ListIndex
    TxtName.Text = List1.Text
End Sub
Private Sub mnuCopy_Click()
' Copy entry
    On Error Resume Next
    EntryMemo = List1.Text
    mnuPaste.Enabled = True
End Sub
Private Sub mnuPaste_Click()
' Paste entry
    On Error Resume Next
    If EntryMemo <> "" Then
        List1.AddItem EntryMemo, List1.ListIndex + 1
        TxtName.Text = EntryMemo
    End If
End Sub
Private Sub mnuCut_Click()
' Cut entry
    On Error Resume Next
    EntryMemo = List1.Text
    List1.RemoveItem List1.ListIndex
    mnuPaste.Enabled = True
End Sub
Private Sub mnuSortAZ_Click()
' Sort List A->Z (put list1->Templist which is sorted and bring all back in order)
    TempList.Clear
    List1.ListIndex = 0
    On Error Resume Next
    For I = 0 To List1.ListCount - 1 Step 1
        TempList.AddItem List1.List(I)
    Next I
    List1.Clear
    For I = 0 To TempList.ListCount - 1 Step 1
        List1.AddItem TempList.List(I)
    Next I
    TempList.Clear
End Sub
Private Sub mnuSortZA_Click()
' Sort List Z->A (put list1->Templist which is sorted and bring all back in reversed order)
    TempList.Clear
    List1.ListIndex = 0
    On Error Resume Next
    For I = 0 To List1.ListCount - 1 Step 1
        TempList.AddItem List1.List(I)
    Next I
    List1.Clear
    For I = TempList.ListCount To 0 Step -1
        List1.AddItem TempList.List(I)
    Next I
    ' Remove "void" entry bug
    List1.ListIndex = 0
    List1.RemoveItem List1.ListIndex
    TempList.Clear
End Sub
Private Sub mnuWinVol_Click()
' Show Windows Volume Controls - System tray option
    On Error GoTo Problem
    Dim lngresult As Long
    lngresult = Shell("c:\windows\Sndvol32.exe", vbNormalFocus)
    Exit Sub
Problem:
    lngresult = Shell("c:\winnt\system32\Sndvol32.exe", vbNormalFocus)
End Sub
Private Sub mnuWinVol2_Click()
' Show Windows Volume Controls - List option
    mnuWinVol_Click
End Sub
'-----------------------'
'GREQ BUTTON/IMAGE CLICK'
'-----------------------'
Private Sub Image_GrEq_Click()
    frmGrEq.Show 0, Me
End Sub
'--------------------------'
'COLLECTION UTILITY BUTTONS'
'--------------------------'
Private Sub cmdCUtility_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Collection Utility button update
    cmdCUtility.Picture = ImgCUtilityon.Picture
End Sub
Private Sub cmdCUtility_Click()
' Show collections utility
    frmCollection.Show
    RETURN_ADD_DIR
End Sub
Private Sub cmdCUtility_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Collection Utility button update
    cmdCUtility.Picture = ImgCUtilityoff.Picture
End Sub
