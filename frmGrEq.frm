VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGrEq 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   1215
   ClientLeft      =   1965
   ClientTop       =   2310
   ClientWidth     =   2775
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmGrEq.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   81
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   185
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Slider Slider_color 
      Height          =   255
      Left            =   0
      TabIndex        =   5
      ToolTipText     =   "Graphic Equalizer Color Density"
      Top             =   0
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      LargeChange     =   15
      Max             =   255
      SelStart        =   255
      TickStyle       =   3
      Value           =   255
   End
   Begin MSComctlLib.Slider Slider 
      Height          =   255
      Left            =   0
      TabIndex        =   2
      ToolTipText     =   "Graphic Equalizer Sensitivity of Analysis"
      Top             =   960
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Max             =   500
      SelStart        =   500
      TickStyle       =   3
      Value           =   500
   End
   Begin VB.PictureBox Scope 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      FontTransparent =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   71
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   183
      TabIndex        =   0
      Top             =   0
      Width           =   2775
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Color Density"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   180
         Left            =   0
         TabIndex        =   6
         Top             =   240
         Width           =   870
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Sensitivity"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   180
         Left            =   2040
         TabIndex        =   4
         Top             =   765
         Width           =   675
      End
   End
   Begin VB.Timer Timer_Greq 
      Interval        =   1000
      Left            =   1920
      Top             =   240
   End
   Begin VB.PictureBox ScopeBuff 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00404040&
      Height          =   330
      Left            =   1320
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.ComboBox DevicesBox 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Menu mnuGrEqOptions 
      Caption         =   "GrEq Options"
      Visible         =   0   'False
      Begin VB.Menu mnuMaxSen 
         Caption         =   "Max Sensitivity"
      End
      Begin VB.Menu mnuMinSen 
         Caption         =   "Min Sensitivity"
      End
      Begin VB.Menu mnuNothing10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMaxCD 
         Caption         =   "Max Color Density"
      End
      Begin VB.Menu mnuMinCD 
         Caption         =   "Min Color Density"
      End
      Begin VB.Menu mnuNothing11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRed 
         Caption         =   "Red Palette"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuGreen 
         Caption         =   "Green  Palette"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuBlue 
         Caption         =   "Blue Palette"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "frmGrEq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'--------------'
'GREQ FUNCTIONS'
'--------------'
Private Sub Form_Load()
    ' Initialize GrEq Form
    mnuRed.Checked = True
    mnuGreen.Checked = False
    mnuBlue.Checked = False
    GREQ_ENABLED = True
    frmGrEq.Top = frmMain.Top + 120
    frmGrEq.Left = frmMain.Left + 4320
    INITIALIZE_GREQ
End Sub
Private Sub Slider_color_Click()
' Change color attributes
    GREQ_COLOR = Slider_color.Value
End Sub
Private Sub Timer_Greq_Timer()
' Timer for greq update
Static WAVEFORMAT As WaveFormatEx
    With WAVEFORMAT
        .FormatTag = WAVE_FORMAT_PCM
        .Channels = 1
        .SamplesPerSec = 11025 '11khz
        .BitsPerSample = 16
        .BlockAlign = (.Channels * .BitsPerSample) \ 8
        .AvgBytesPerSec = .BlockAlign * .SamplesPerSec
        .ExtraDataSize = 0
    End With
    Debug.Print "waveInOpen:"; waveInOpen(DevHandle, DevicesBox.ListIndex, VarPtr(WAVEFORMAT), 0, 0, 0)
    If DevHandle = 0 Then
        Call MsgBox("Wave input device didn't open and the" & vbNewLine & "Graphic Equalizer will not function...", vbExclamation, "Graphic Equalizer")
        Timer_Greq.Enabled = False
        Exit Sub
    End If
    Debug.Print " "; DevHandle
    Call waveInStart(DevHandle)
    Slider.Enabled = True
    DevicesBox.Enabled = False
    Timer_Greq.Enabled = False
    Call Visualize
End Sub
Public Sub INITIALIZE_GREQ()
' Initialize graphic equalizer
    Call InitDevices
    Call DoReverse
    Call Slider_Change
    ScopeBuff.Width = Scope.ScaleWidth
    ScopeBuff.Height = Scope.ScaleHeight
    ScopeBuff.BackColor = Scope.BackColor
    ScopeHeight = Scope.Height
    Timer_Greq.Enabled = True
    On Error Resume Next
End Sub
Public Sub Slider_Change()
' Change the greq divisor
    Divisor = ((Slider.Max - Slider.Value + 1) / Slider.Max) * 5200
End Sub
Public Sub InitDevices()
' Initialize sound device for greq
    Dim Caps As WAVEINCAPS, Which As Long
    DevicesBox.Clear
    For Which = 0 To waveInGetNumDevs - 1
        Call waveInGetDevCaps(Which, VarPtr(Caps), Len(Caps))
    Next
End Sub
Public Sub Visualize()
' Greq working
    On Error Resume Next
    Visualizing = True
    Static X As Long
    Static Wave As WAVEHDR
    Static InData(0 To NUMSAMPLES - 1) As Integer
    Static OutData(0 To NUMSAMPLES - 1) As Single
    With ScopeBuff
        Do
            Wave.lpData = VarPtr(InData(0))
            Wave.dwBufferLength = NUMSAMPLES
            Wave.dwFlags = 0
            Call waveInPrepareHeader(DevHandle, VarPtr(Wave), Len(Wave))
            Call waveInAddBuffer(DevHandle, VarPtr(Wave), Len(Wave))
            Do
            DoEvents
            Loop Until ((Wave.dwFlags And WHDR_DONE) = WHDR_DONE) Or DevHandle = 0
            If DevHandle = 0 Then Exit Do
            Call waveInUnprepareHeader(DevHandle, VarPtr(Wave), Len(Wave))
            Call FFTAudio(InData, OutData)
            .Cls
            .CurrentX = 0
            .CurrentY = ScopeHeight
            For X = 0 To 255 Step 2
                .CurrentY = ScopeHeight
                .CurrentX = X
                If mnuRed.Checked = True Then
                    ' RED
                    ScopeBuff.ForeColor = RGB(GREQ_COLOR, X, 0)
                Else
                    If mnuGreen.Checked = True Then
                        ' GREEN
                        ScopeBuff.ForeColor = RGB(0, GREQ_COLOR, X)
                    Else
                        ' BLUE
                        ScopeBuff.ForeColor = RGB(X, 0, GREQ_COLOR)
                    End If
                End If
                ScopeBuff.Line Step(0, 0)-(X, ScopeHeight - (Sqr(Abs(OutData(X * 2) \ Divisor)) / 2 + Sqr(Abs(OutData(X * 2 + 1) \ Divisor)) / 2))
            Next X
            Scope.Picture = .Image
            DoEvents
        Loop While DevHandle <> 0
    End With
    Visualizing = False
End Sub
'--------------'
'MENU SHORTCUTS'
'--------------'
Private Sub mnuBlue_Click()
    mnuRed.Checked = False
    mnuGreen.Checked = False
    mnuBlue.Checked = True
End Sub
Private Sub mnuGreen_Click()
    mnuRed.Checked = False
    mnuGreen.Checked = True
    mnuBlue.Checked = False
End Sub
Private Sub mnuRed_Click()
    mnuRed.Checked = True
    mnuGreen.Checked = False
    mnuBlue.Checked = False
End Sub
Private Sub Scope_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Show GrEq pop up menu
    If Button = vbRightButton Then
        PopupMenu Me.mnuGrEqOptions
    End If
End Sub
Private Sub mnuMaxCD_Click()
' Maximum color density
    Slider_color.Value = 255
    Slider_color_Click
End Sub
Private Sub mnuMinCD_Click()
' Minimum color density
    Slider_color.Value = 0
    Slider_color_Click
End Sub
Private Sub mnuMaxSen_Click()
' Maximum Sensitivity
    Slider.Value = 500
End Sub
Private Sub mnuMinSen_Click()
' Minimum Sensitivity
    Slider.Value = 0
End Sub
