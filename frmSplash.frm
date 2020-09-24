VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Rasputin Player"
   ClientHeight    =   1425
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   3825
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   3825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Picture         =   "frmSplash.frx":030A
         Top             =   360
         Width           =   480
      End
      Begin VB.Label lblCompany 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Zamora Software"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2040
         TabIndex        =   7
         Top             =   1080
         Width           =   1485
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Version 1.XXXX.XXXX"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   210
         Left            =   1800
         TabIndex        =   6
         Top             =   840
         Width           =   1740
      End
      Begin VB.Label lblmail 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Zozas@hotmail.com"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   1470
      End
      Begin VB.Label lblMe 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Zozas Ioannis"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   1140
      End
      Begin VB.Label lblDescription 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Audio Player for Windows 9x/Me and NT/2000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   3375
      End
      Begin VB.Label lblName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Rasputin"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1335
         TabIndex        =   2
         Top             =   360
         Width           =   945
      End
      Begin VB.Label lblPlayer 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Player    "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2280
         TabIndex        =   1
         Top             =   480
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'--------------------------------------------'
'LOADING PROGRAM/STARTING VALUES/INITIALIZING'
'--------------------------------------------'
Private Sub Form_Load()
' Initializing forms and giving starting values to variables
    ' Detect sound cards
    DETECT_SOUND_CARD
    ' Morph splash screen
    lblName.Caption = App.ProductName
    lblPlayer.Caption = "Player    "
    lblMe.Caption = App.LegalTrademarks
    lblCompany.Caption = App.CompanyName
    lblmail.Caption = App.LegalCopyright
    lblDescription.Caption = "Audio Player for Windows 9x/Me and NT/2000"
    LblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    ' Check program options for variables
    frmSplash.Show
    ' Normalize main form
    frmMain.Width = 4215
    frmMain.Height = 6630
    FORM_MODE = True
    LOAD_OPTIONS
    Set PC_FILE_SYSTEM = CreateObject("Scripting.FileSystemObject")
    WINDOWS_PATH = PC_FILE_SYSTEM.getspecialfolder(0)
    MINUTES_LEFT = "00"
    SECONDS_LEFT = "00"
    volume = 2500
    IS_MINIMIZED = False
    PAUSED_PLAYER = False
    MOVE_FORM = False
    MOVE_OPTIONS = False
    '
    '
    '
    GREQ_ENABLED = False
    GREQ_COLOR = 255
    '
    '
    FORM_TITLE = "Rasputin - No Title *** "
    SysTrayTitle = FORM_TITLE
    ' Initializing Elements
    frmMain.sldVol.Value = frmMain.MediaPlayer1.volume
    frmMain.sldVol.Value = 2500
    frmMain.LabelVolume.ForeColor = RGB(155 + frmMain.sldVol.Value / 10, 100, 100)
    frmMain.LabelVolume.Caption = "100 %"
    frmMain.ListInfo.Text = "0/0 File"
    frmMain.cmdCompact.ToolTipText = "Compact mode"
    ' Preparing form
    SEQUENCE
    RETURN_ADD_DIR
    AUTOLOAD_PLAYLIST
    PREPARE_MAIN_FORM
'
'
'
'
'
'
    
    ' Loading Program
    frmMain.TimerTitle.Interval = TIMER_SPEED
    frmMain.Show
    Unload Me
End Sub
