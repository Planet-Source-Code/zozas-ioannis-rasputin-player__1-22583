VERSION 5.00
Begin VB.Form frmCollection 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Collection Manager"
   ClientHeight    =   3675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3630
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   3630
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3450
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   3375
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   2550
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label lblInfo2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "(Select and Delete key to remove file)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   3135
      End
      Begin VB.Image cmdAdd 
         Height          =   225
         Left            =   720
         Picture         =   "frmCollection.frx":0000
         ToolTipText     =   "Add sound files above to the playlist"
         Top             =   3120
         Width           =   660
      End
      Begin VB.Label lbl_files 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Sound Files found in Subdirectories"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   210
         Left            =   120
         TabIndex        =   10
         Top             =   0
         Width           =   3060
      End
      Begin VB.Image cmdCancel 
         Height          =   225
         Index           =   2
         Left            =   1920
         Picture         =   "frmCollection.frx":07FE
         ToolTipText     =   "Cancel action and return"
         Top             =   3120
         Width           =   660
      End
      Begin VB.Image cmdGoBack 
         Height          =   225
         Index           =   1
         Left            =   1320
         Picture         =   "frmCollection.frx":0FFC
         ToolTipText     =   "Go back to sub folder results"
         Top             =   3120
         Width           =   660
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3450
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3375
      Begin VB.ListBox List_SubDirs 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   2550
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label lblInfo1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(Select and Delete key to remove directory)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   3135
      End
      Begin VB.Image cmdScan 
         Height          =   225
         Left            =   720
         Picture         =   "frmCollection.frx":17FA
         ToolTipText     =   "Scan folders above for sound files"
         Top             =   3120
         Width           =   660
      End
      Begin VB.Image cmdGoBack 
         Height          =   225
         Index           =   0
         Left            =   1320
         Picture         =   "frmCollection.frx":1FF8
         ToolTipText     =   "Go back to folder selection"
         Top             =   3120
         Width           =   660
      End
      Begin VB.Image cmdCancel 
         Height          =   225
         Index           =   1
         Left            =   1920
         Picture         =   "frmCollection.frx":27F6
         ToolTipText     =   "Cancel action and return"
         Top             =   3120
         Width           =   660
      End
      Begin VB.Label lblFound 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Directories Found"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   0
         Width           =   3135
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3450
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3375
      Begin VB.DirListBox Dir1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   2250
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   3135
      End
      Begin VB.DriveListBox Drive1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   330
         Left            =   120
         TabIndex        =   2
         Top             =   3030
         Width           =   3135
      End
      Begin VB.Label lblInfo0 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Please Wait while scanning..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   3135
      End
      Begin VB.Image cmdLup 
         Height          =   225
         Left            =   1920
         Picture         =   "frmCollection.frx":2FF4
         Stretch         =   -1  'True
         ToolTipText     =   "Directory level up"
         Top             =   2760
         Width           =   660
      End
      Begin VB.Image cmdDesktop 
         Height          =   225
         Left            =   1320
         Picture         =   "frmCollection.frx":37F2
         Stretch         =   -1  'True
         ToolTipText     =   "Go to desktop"
         Top             =   2760
         Width           =   660
      End
      Begin VB.Image cmdMyDocuments 
         Height          =   225
         Left            =   720
         Picture         =   "frmCollection.frx":3FF0
         Stretch         =   -1  'True
         ToolTipText     =   "Go to my documents"
         Top             =   2760
         Width           =   660
      End
      Begin VB.Label lblDir 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Directory to search"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   0
         Width           =   3135
      End
      Begin VB.Image cmdSSub 
         Height          =   225
         Left            =   120
         Picture         =   "frmCollection.frx":47EE
         ToolTipText     =   "Scan to find all subdirectories"
         Top             =   2760
         Width           =   660
      End
      Begin VB.Image cmdCancel 
         Height          =   225
         Index           =   0
         Left            =   2520
         Picture         =   "frmCollection.frx":4FEC
         ToolTipText     =   "Cancel and return"
         Top             =   2760
         Width           =   660
      End
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   225
      Left            =   120
      Pattern         =   "*.mp3;*.wav"
      TabIndex        =   0
      Top             =   3600
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "frmCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private DirName As String       ' Directory-to-scan name
'---------------------------------'
'LOAD FORM AND INITIALIZE ELEMENTS'
'---------------------------------'
Private Sub Form_Load()
    MOVE_SCAN = False
    lblInfo0.Visible = False
    Frame3.Visible = False
    Frame2.Visible = False
    Frame1.Visible = True
    frmCollection.Picture = frmMain.Picture
End Sub
Private Sub Dir1_Change()
' Change file path
    File1.Path = Dir1.Path
End Sub
Private Sub Drive1_Change()
' Change dir path
    Dir1.Path = Drive1.Drive
End Sub
'-------------'
'FORM MOVEMENT'
'-------------'
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Start of form movement
    MOVE_SCAN = True
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Move form if key mouse pressed
    If MOVE_SCAN = True Then
        frmCollection.Move frmCollection.Left + (X - PREVIOUS_SCAN_X), frmCollection.Top + (Y - PREVIOUS_SCAN_Y)
    Else
        PREVIOUS_SCAN_X = X
        PREVIOUS_SCAN_Y = Y
    End If
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' End of form movement
    MOVE_SCAN = False
End Sub
'-------------'
'LIST HANDLING'
'-------------'
Private Sub List_SubDirs_KeyDown(KeyCode As Integer, Shift As Integer)
' Remove if Delete key is pressed
    If KeyCode = vbKeyDelete Then
        List_SubDirs.RemoveItem List_SubDirs.ListIndex
    End If
End Sub
Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
' Remove if Delete key is pressed
    If KeyCode = vbKeyDelete Then
        frmCollection.List1.RemoveItem frmCollection.List1.ListIndex
    End If
End Sub '------------------'
'CANCEL SCAN ACTION'
'------------------'
Private Sub cmdCancel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
' Cancel button update
    cmdCancel(Index).Picture = frmMain.ImgCancelon.Picture
End Sub
Private Sub cmdCancel_Click(Index As Integer)
' Cancel action and return
    Unload Me
End Sub
Private Sub cmdCancel_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
' Cancel button update
    cmdCancel(Index).Picture = frmMain.ImgCanceloff.Picture
End Sub
'--------------------'
'GO TO DESKTOP FOLDER'
'--------------------'
Private Sub cmdDesktop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Go to desktop button update
    cmdDesktop.Picture = frmMain.ImgDeskon.Picture
End Sub
Private Sub cmdDesktop_Click()
' Go to desktop
    On Error Resume Next
    Dir1.Path = WINDOWS_PATH & "\Desktop"
    Drive1.Drive = Left(WINDOWS_PATH, 1)
End Sub
Private Sub cmdDesktop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Go to desktop button update
    cmdDesktop.Picture = frmMain.ImgDeskoff.Picture
End Sub
'------------------------------'
'SCAN FOR SUBDIRECTORIES BUTTON'
'------------------------------'
Private Sub cmdSSub_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Scan Subdirectories button update
    cmdSSub.Picture = frmMain.ImgScanSubs_on.Picture
End Sub
Private Sub cmdSSub_Click()
' Scan for all subdirectories
    lblInfo0.Visible = True
    cmdSSub.Enabled = False
    List_SubDirs.Clear
    DirName = Dir1.Path
    GoDeep DirName, 0
    Frame1.Visible = False
    Frame3.Visible = False
    Frame2.Visible = True
    cmdSSub.Enabled = True
    lblInfo0.Visible = False
End Sub
Private Sub cmdSSub_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Scan Subdirectories button update
    cmdSSub.Picture = frmMain.ImgScanSubs_off.Picture
End Sub
Public Function GoDeep(ByVal Whereto As String, ByVal dircount As Integer)
    Dim Subs() As String
    Dim Count As Integer
    Dim Plekk As String
    Dim I As Integer
    If Right(Whereto, 1) <> "\" Then
        Whereto = Whereto & "\"
    End If
    Count = 0
    Plekk = Dir(Whereto, vbDirectory)
    Do While Plekk <> ""
        DoEvents
        If (Plekk = ".") Or (Plekk = "..") Then
        Else
            If GetAttr(Whereto & Plekk) = vbDirectory Then
                ReDim Preserve Subs(Count)
                Subs(Count) = Plekk
                Count = Count + 1
                frmCollection.List_SubDirs.AddItem Whereto & Plekk
            End If
        End If
        Plekk = Dir()
    Loop
    If Count > 0 Then
        For I = 0 To Count - 1
            GoDeep Whereto & Subs(I) & "\", Count
        Next
    End If
End Function
'--------------'
'GO BACK BUTTON'
'--------------'
Private Sub cmdGoBack_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
' Go Back button update
    cmdGoBack(Index).Picture = frmMain.ImgBack_on.Picture
End Sub
Private Sub cmdGoBack_Click(Index As Integer)
' Go back
Select Case Index
Case Is = 0
    Frame2.Visible = False
    Frame3.Visible = False
    Frame1.Visible = True
Case Is = 1
    Frame3.Visible = False
    Frame1.Visible = False
    Frame2.Visible = True
End Select
End Sub
Private Sub cmdGoBack_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
' Go Back button update
    cmdGoBack(Index).Picture = frmMain.ImgBack_Off.Picture
End Sub
'-------------------------'
'GO TO MY DOCUMENTS FOLDER'
'-------------------------'
Private Sub cmdMyDocuments_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Go to My documents button update
    cmdMyDocuments.Picture = frmMain.ImgMyon.Picture
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
' Go to My documents button update
    cmdMyDocuments.Picture = frmMain.ImgMyoff.Picture
End Sub
'------------------'
'DIRECTORY LEVEL UP'
'------------------'
Private Sub cmdLup_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Directory Level Up in the Dir List Box
    cmdLup.Picture = frmMain.ImgLupon.Picture
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
' Directory Level Up in the Dir List Box
    cmdLup.Picture = frmMain.ImgLupoff.Picture
End Sub
'----------------------------'
'SCAN FOLDERS FOR SOUND FILES'
'----------------------------'
Private Sub cmdScan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Scan button update
    cmdScan.Picture = frmMain.ImgScanon.Picture
End Sub
Private Sub cmdScan_Click()
' Scan folders for sound files
Dim I As Integer
Dim j As Integer
    For I = 0 To List_SubDirs.ListCount - 1 Step 1
        List_SubDirs.ListIndex = I
        Dir1.Path = List_SubDirs.Text
        For j = 0 To File1.ListCount - 1
            File1.ListIndex = j
            List1.AddItem (Dir1.Path & "\" & File1.FileName)
        Next j
    Next I
    Frame1.Visible = False
    Frame2.Visible = False
    Frame3.Visible = True
End Sub
Private Sub cmdScan_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Scan button update
    cmdScan.Picture = frmMain.ImgScanoff.Picture
End Sub
'---------------------------'
'ADD FILES FOUND TO PLAYLIST'
'---------------------------'
Private Sub cmdAdd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Button Add update
    cmdAdd.Picture = frmMain.ImgConfirmon.Picture
End Sub
Private Sub cmdAdd_Click()
' Add files to playlist
    lblInfo0.Visible = False
    lblInfo2.Caption = lblInfo0.Caption
    On Error Resume Next
    For I = 0 To frmCollection.List1.ListCount - 1
        frmCollection.List1.ListIndex = I
        frmMain.List1.AddItem frmCollection.List1.Text
    Next I
    lblInfo2.Caption = "(Select and Delete key to remove file)"
    Unload Me
End Sub
Private Sub cmdAdd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Button Add update
    cmdAdd.Picture = frmMain.ImgConfirmoff.Picture
End Sub
