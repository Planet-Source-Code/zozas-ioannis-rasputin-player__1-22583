VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Options"
   ClientHeight    =   4290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3615
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   25
      Left            =   3240
      Top             =   3960
   End
   Begin VB.Frame FrameOptions 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   3735
      Index           =   0
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   3015
      Begin VB.TextBox TxtListPath 
         ForeColor       =   &H00008000&
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Text            =   "Filename"
         ToolTipText     =   "Auto List filename"
         Top             =   2760
         Width           =   2655
      End
      Begin VB.CheckBox chkAutoList 
         BackColor       =   &H00000000&
         Caption         =   "Use Auto List"
         ForeColor       =   &H00C0C000&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "Save current playlist when quitting and reload it on startup"
         Top             =   2400
         Width           =   1335
      End
      Begin VB.CheckBox chkAutoplay 
         BackColor       =   &H00000000&
         Caption         =   "Autoplay Song/Track"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Play song at startup"
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox TxtFF 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   120
         MaxLength       =   2
         TabIndex        =   6
         Text            =   "00"
         ToolTipText     =   "Fast Forward seconds"
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox TxtRW 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   120
         MaxLength       =   2
         TabIndex        =   5
         Text            =   "00"
         ToolTipText     =   "Rewind seconds"
         Top             =   1440
         Width           =   375
      End
      Begin VB.CheckBox OptNone 
         BackColor       =   &H00000000&
         Caption         =   "Don't play a track as Next"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "After playing a track, don't play any other tracks"
         Top             =   240
         Width           =   2175
      End
      Begin VB.CheckBox OptRnd 
         BackColor       =   &H00000000&
         Caption         =   "Play Random track as Next"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "After playing a track, play a random track from the playlist"
         Top             =   720
         Width           =   2295
      End
      Begin VB.CheckBox OptNext 
         BackColor       =   &H00000000&
         Caption         =   "Play Next track as Next"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "After playing a track, play the next track after it"
         Top             =   480
         Width           =   2055
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   2280
         Top             =   2160
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label cmdLookUp 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Auto List..."
         ForeColor       =   &H00FFFF80&
         Height          =   195
         Left            =   2040
         TabIndex        =   9
         ToolTipText     =   "Choose Autolist file"
         Top             =   3060
         Width           =   750
      End
      Begin VB.Label LblFF 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Set Fast Forward (1-99 seconds)"
         ForeColor       =   &H0080C0FF&
         Height          =   195
         Left            =   600
         TabIndex        =   22
         Top             =   1800
         Width           =   2295
      End
      Begin VB.Label LblRW 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Set Rewind (1-99 seconds)"
         ForeColor       =   &H0080C0FF&
         Height          =   195
         Left            =   600
         TabIndex        =   21
         Top             =   1440
         Width           =   1920
      End
   End
   Begin VB.Frame FrameOptions 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3735
      Index           =   1
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   3015
      Begin MSComctlLib.Slider sldTS 
         Height          =   255
         Left            =   1440
         TabIndex        =   17
         ToolTipText     =   "Define the speed to scroll the title"
         Top             =   2280
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         LargeChange     =   2
      End
      Begin VB.CheckBox ChkonTop 
         BackColor       =   &H00000000&
         Caption         =   "Keep Program on Top"
         ForeColor       =   &H00C0C0FF&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         ToolTipText     =   "Keep program on top of all other programs"
         Top             =   1680
         Width           =   2415
      End
      Begin VB.CheckBox ChkTitleScroll 
         BackColor       =   &H00000000&
         Caption         =   "Scroll Song Title on Taskbar"
         ForeColor       =   &H00C0C0FF&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         ToolTipText     =   "Scroll song title on taskbar"
         Top             =   1920
         Width           =   2415
      End
      Begin VB.CheckBox OptCompact 
         BackColor       =   &H00000000&
         Caption         =   "Compact mode on Startup"
         ForeColor       =   &H0080FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   2295
      End
      Begin VB.CheckBox OptTools 
         BackColor       =   &H00000000&
         Caption         =   "Show List/Tools on Startup"
         ForeColor       =   &H0080FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   2295
      End
      Begin VB.CheckBox OptList 
         BackColor       =   &H00000000&
         Caption         =   "Show List on Startup"
         ForeColor       =   &H0080FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1815
      End
      Begin VB.CheckBox chkMinSys 
         BackColor       =   &H00000000&
         Caption         =   "Minimize on Taskbar"
         ForeColor       =   &H0080FF80&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         ToolTipText     =   "Minimize on Taskbar instead of System Tray"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.CheckBox chkSysMin 
         BackColor       =   &H00000000&
         Caption         =   "Minimize on System Tray"
         ForeColor       =   &H0080FF80&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         ToolTipText     =   "Minimize on System Tray instead of Taskbar"
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label lblTspeed 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Title Scroll Speed"
         ForeColor       =   &H00FFFFC0&
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   2280
         Width           =   1245
      End
   End
   Begin VB.Frame FrameOptions 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3735
      Index           =   2
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   3015
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   2895
         Left            =   90
         ScaleHeight     =   2865
         ScaleWidth      =   2805
         TabIndex        =   26
         Top             =   720
         Width           =   2830
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   $"frmOptions.frx":030A
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   8895
            Left            =   0
            TabIndex        =   27
            Top             =   2760
            Width           =   2775
         End
      End
      Begin VB.Label lblWinpath 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Windows Path : C:\WINDOWS"
         ForeColor       =   &H00C0FFC0&
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label LblVersion 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Rasputin Player Version 1.0.115"
         ForeColor       =   &H00C0FFC0&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   2775
      End
   End
   Begin MSComctlLib.TabStrip TabOptions 
      Height          =   3735
      Left            =   3120
      TabIndex        =   0
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   6588
      MultiRow        =   -1  'True
      HotTracking     =   -1  'True
      Placement       =   3
      Separators      =   -1  'True
      TabStyle        =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Play Mode"
            Object.ToolTipText     =   "Play mode and handling options"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Visualization"
            Object.ToolTipText     =   "Vizualization options"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "About Program"
            Object.ToolTipText     =   "Credits and information about Rasputin Player"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image cmdApply 
      Height          =   225
      Left            =   1200
      Picture         =   "frmOptions.frx":0524
      ToolTipText     =   "Apply current values"
      Top             =   3960
      Width           =   660
   End
   Begin VB.Image cmdOK 
      Height          =   225
      Left            =   600
      Picture         =   "frmOptions.frx":0D22
      ToolTipText     =   "Apply changes and return"
      Top             =   3960
      Width           =   660
   End
   Begin VB.Image cmdCancel 
      Height          =   225
      Left            =   2400
      Picture         =   "frmOptions.frx":1520
      ToolTipText     =   "Cancel all changes and return"
      Top             =   3960
      Width           =   660
   End
   Begin VB.Image cmdDefault 
      Height          =   225
      Left            =   1800
      Picture         =   "frmOptions.frx":1D1E
      ToolTipText     =   "Show default values"
      Top             =   3960
      Width           =   660
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mintCurFrame As Integer                 'Tab index
'-------------------------'
'      APPLY CHANGES      '
'-------------------------'
Private Sub cmdApply_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Apply button update
    cmdApply.Picture = frmMain.ImgApplyOn.Picture
End Sub
Private Sub cmdApply_Click()
' Apply chancges and write options file
    SaveSetting App.Title, "Startup", "RW", TxtRW.Text
    SaveSetting App.Title, "Startup", "FF", TxtFF.Text
    ' Next song
    If OptNext.Value = 1 Then
        SaveSetting App.Title, "Startup", "NEXT_SONG", "1"
    Else
        If OptRnd.Value = 1 Then
            SaveSetting App.Title, "Startup", "NEXT_SONG", "2"
        Else
            SaveSetting App.Title, "Startup", "NEXT_SONG", "3"
        End If
    End If
    ' Form mode
    If OptList.Value = 1 Then
        SaveSetting App.Title, "Startup", "LIST", "1"
    Else
        SaveSetting App.Title, "Startup", "LIST", "0"
    End If
    If OptTools.Value = 1 Then
        SaveSetting App.Title, "Startup", "TOOLS", "1"
    Else
        SaveSetting App.Title, "Startup", "TOOLS", "0"
    End If
    If OptCompact.Value = 1 Then
        SaveSetting App.Title, "Startup", "COMPACT", "1"
    Else
        SaveSetting App.Title, "Startup", "COMPACT", "0"
    End If
    ' On top
    If ChkonTop.Value = 1 Then
        SaveSetting App.Title, "Startup", "ON_TOP", "1"
    Else
        SaveSetting App.Title, "Startup", "ON_TOP", "0"
    End If
    ' Scroll Title
    If ChkTitleScroll.Value = 1 Then
        SaveSetting App.Title, "Startup", "TITLE_SCROLL", "1"
    Else
        SaveSetting App.Title, "Startup", "TITLE_SCROLL", "0"
    End If
    ' Autolist
    If chkAutoList.Value = 1 Then
        SaveSetting App.Title, "Startup", "AUTO_LIST", "1"
    Else
        SaveSetting App.Title, "Startup", "AUTO_LIST", "0"
    End If
    ' Minimize mode
    If chkSysMin.Value = 1 Then
        SaveSetting App.Title, "Startup", "MINIMIZE_MODE", "1"
    Else
        SaveSetting App.Title, "Startup", "MINIMIZE_MODE", "0"
    End If
    ' Autoplay
    If chkAutoplay.Value = 1 Then
        SaveSetting App.Title, "Startup", "AUTOPLAY", "1"
    Else
        SaveSetting App.Title, "Startup", "AUTOPLAY", "0"
    End If
    ' Autolist file
    SaveSetting App.Title, "Startup", "AUTO_LIST_FILE", TxtListPath.Text
    SaveSetting App.Title, "Startup", "TIMER_SPEED", (sldTS.Value * 50 + 150)
End Sub
Private Sub cmdApply_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Apply button update
    cmdApply.Picture = frmMain.ImgApplyoff.Picture
End Sub
'-------------------------'
'QUIT FORM WITHOUT SAVING '
'-------------------------'
Private Sub cmdDefault_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Default button update
    cmdDefault.Picture = frmMain.ImgOptDefOn.Picture
End Sub
Private Sub cmdDefault_Click()
    ' Set default options
    OptCompact.Value = 0
    OptTools.Value = 0
    OptList.Value = 0
    OptNext.Value = 0
    OptRnd.Value = 0
    OptNone.Value = 0
    TxtFF.Text = DEFAULT_FF
    TxtRW.Text = DEFAULT_RW
    If chkAutoList.Value = 0 Then
        cmdLookUp.Enabled = False
        TxtListPath.Enabled = False
    End If
    Select Case DEFAULT_NEXT_SONG
    Case Is = 1
        OptNext.Value = 1
    Case Is = 2
        OptRnd.Value = 1
    Case Is = 3
        OptNone.Value = 1
    End Select
    If DEFAULT_FORM_MODE = True Then
        OptCompact.Value = 1
    Else
        If DEFAULT_FORM_TOOLS = True Then
            OptTools.Value = 1
        Else
            OptList.Value = 1
        End If
    End If
    If DEFAULT_FORM_ON_TOP = True Then
        ChkonTop.Value = 1
    Else
        ChkonTop.Value = 0
    End If
    If DEFAULT_TITLE_SCROLL = True Then
        ChkTitleScroll.Value = 1
    Else
        ChkTitleScroll.Value = 0
    End If
    If DEFAULT_AUTOLOAD_LIST = True Then
        chkAutoList.Value = 1
    Else
        chkAutoList.Value = 0
    End If
    
    If MINIMIZE_MODE = True Then
        chkSysMin.Value = 1
    Else
        chkMinSys.Value = 0
    End If
    
    TxtListPath.Text = App.Path & "\AutoList.rpl"
End Sub
Private Sub cmdDefault_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Default button update
    cmdDefault.Picture = frmMain.ImgOptDefoff.Picture
End Sub
'--------------------------'
'QUIT FORM AND SAVE CHANGES'
'--------------------------'
Private Sub cmdOK_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' OK button update
    cmdOK.Picture = frmMain.ImgConfirmon.Picture
End Sub
Private Sub cmdOK_Click()
    cmdApply_Click
    Unload Me
End Sub
Private Sub cmdOK_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' OK button update
    cmdOK.Picture = frmMain.ImgConfirmoff.Picture
End Sub
'-------------------------'
'QUIT FORM WITHOUT SAVING '
'-------------------------'
Private Sub cmdCancel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Cancel buton update
    cmdCancel.Picture = frmMain.ImgCancelon.Picture
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdCancel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Cancel buton update
    cmdCancel.Picture = frmMain.ImgCanceloff.Picture
End Sub
'-------------------------'
'LOADING FORM/INITIALIZING'
'-------------------------'
Private Sub Form_Load()
    Timer1.Enabled = False
    FrameOptions(0).Visible = True
    FrameOptions(1).Visible = False
    FrameOptions(2).Visible = False
    TabOptions.TabIndex = 1
    mintCurFrame = 1
    LblVersion.Caption = "Rasputin Player Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblWinpath.Caption = "Windows Path : " & WINDOWS_PATH
    frmOptions.Picture = frmMain.Picture
    TxtFF.Text = FF
    TxtRW.Text = RW
    TxtListPath.Text = AUTOLIST
    OptCompact.Value = 0
    OptTools.Value = 0
    OptList.Value = 0
    OptNext.Value = 0
    OptRnd.Value = 0
    OptNone.Value = 0
    On Error Resume Next
    sldTS.Value = Int((TIMER_SPEED - 150) / 50)
    ' Minimize mode
    If MINIMIZE_MODE = True Then
        chkSysMin.Value = 1
        chkMinSys.Value = 0
    Else
        chkMinSys.Value = 1
        chkSysMin.Value = 0
    End If
    ' Form on top check
    If FORM_ON_TOP = True Then
        ChkonTop.Value = 1
    Else
        ChkonTop.Value = 0
    End If
    ' Autoload list check
    If AUTOLOAD_LIST = True Then
        chkAutoList.Value = 1
    Else
        chkAutoList.Value = 0
        cmdLookUp.Enabled = False
        TxtListPath.Enabled = False
    End If
    ' Title scroll check
    If TITLE_SCROLL = True Then
        ChkTitleScroll.Value = 1
        lblTspeed.Enabled = True
        sldTS.Enabled = True
    Else
        ChkTitleScroll.Value = 0
        lblTspeed.Enabled = False
        sldTS.Enabled = False
    End If
    ' Mode check
    If FORM_MODE = True Then
        OptCompact.Value = 1
    Else
        If FORM_TOOLS = True Then
            OptTools.Value = 1
        Else
            OptList.Value = 1
        End If
    End If
    ' Next song check
    Select Case NEXT_SONG
    Case Is = 1
        OptNext.Value = 1
    Case Is = 2
        OptRnd.Value = 1
    Case Is = 3
        OptNone.Value = 1
    End Select
    ' aUTOPLAY CHECK
    Select Case AUTOPLAY
    Case Is = True
        chkAutoplay.Value = 1
    Case Is = False
        chkAutoplay.Value = 0
    End Select
End Sub
'-------------------------'
'REACT TO CHANGES ON FORM '
'-------------------------'
Private Sub ChkTitleScroll_Click()
    If ChkTitleScroll.Value = 0 Then
        lblTspeed.Enabled = False
        sldTS.Enabled = False
    Else
        lblTspeed.Enabled = True
        sldTS.Enabled = True
    End If
End Sub
Private Sub chkMinSys_Click()
    chkSysMin.Value = 0
End Sub
Private Sub chkSysMin_Click()
    chkMinSys.Value = 0
End Sub
Private Sub OptCompact_Click()
    OptList.Value = 0
    OptTools.Value = 0
End Sub
Private Sub OptList_Click()
    OptTools.Value = 0
    OptCompact.Value = 0
End Sub
Private Sub OptTools_Click()
    OptList.Value = 0
    OptCompact.Value = 0
End Sub
Private Sub OptNext_Click()
    OptRnd.Value = 0
    OptNone.Value = 0
End Sub
Private Sub OptRnd_Click()
    OptNext.Value = 0
    OptNone.Value = 0
End Sub
Private Sub OptNone_Click()
    OptNext.Value = 0
    OptRnd.Value = 0
End Sub
Private Sub chkAutoList_Click()
    If chkAutoList.Value = 0 Then
        cmdLookUp.Enabled = False
        TxtListPath.Enabled = False
    Else
        cmdLookUp.Enabled = True
        TxtListPath.Enabled = True
    End If
End Sub
'-------------------------'
'    MOVE OPTIONS FORM    '
'-------------------------'
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Realise the moving will by pressing a mouse key
    MOVE_OPTIONS = True
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Move form only if key pressed on it
    If MOVE_OPTIONS = True Then
        frmOptions.Move frmOptions.Left + (X - PREVIOUS_OPTIONS_X), frmOptions.Top + (Y - PREVIOUS_OPTIONS_Y)
    Else
        PREVIOUS_OPTIONS_X = X
        PREVIOUS_OPTIONS_Y = Y
    End If
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Stop moving will
    MOVE_OPTIONS = False
End Sub
'-------------------------'
'LET USER CHOOSE AUTOLIST '
'-------------------------'
Private Sub cmdLookUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Autolist button update
    cmdLookUp.ForeColor = vbRed
End Sub
Private Sub cmdLookUp_Click()
    ' Choose AutoList file via Common Dialogue
    CommonDialog1.DialogTitle = "Set AutoList File"
    CommonDialog1.MaxFileSize = 16384
    CommonDialog1.FileName = ""
    CommonDialog1.Filter = "Rasputin Playlists (*.rpl)|*.rpl"
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName <> "" Then
        TxtListPath.Text = CommonDialog1.FileName
    End If
End Sub
Private Sub cmdLookUp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Autolist button update
    cmdLookUp.ForeColor = vbBlue
End Sub
'-------------------------'
'   TABSTRIP OPERATIONS   '
'-------------------------'
Private Sub TabOptions_Click()
' Handle tabstrip frames
    Timer1.Enabled = False
    Select Case (TabOptions.SelectedItem.Index - 1)
    Case Is = 0
        FrameOptions(0).Visible = True
        FrameOptions(1).Visible = False
        FrameOptions(2).Visible = False
    Case Is = 1
        FrameOptions(0).Visible = False
        FrameOptions(1).Visible = True
        FrameOptions(2).Visible = False
    Case Is = 2
        Timer1.Enabled = True
        FrameOptions(0).Visible = False
        FrameOptions(1).Visible = False
        FrameOptions(2).Visible = True
    End Select
End Sub










Private Sub Timer1_Timer()

'the code below makes the label2 scroll...
If Label2.Top < Picture1.Height - Picture1.Height - Label2.Height Then
    Label2.Top = Picture1.Height - 1
    
    Label2.Top = Label2.Top - 5
    
Else
    Label2.Top = Label2.Top - 10
    
End If
End Sub

