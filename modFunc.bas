Attribute VB_Name = "modFunc"
Option Explicit
'-------------------------'
'   SPLIT TIME DURATION   '
'-------------------------'
Public Sub SPLIT_MINUTES_SECONDS(MPDuration)
' Split minutes and secods from duration
    MINUTES_LEFT = "00"
    SECONDS_LEFT = "00"
    TempVar = 0
    While MPDuration > 60
        MPDuration = MPDuration - 60
        TempVar = TempVar + 1
    Wend
    MINUTES_LEFT = TempVar
    If Sgn(MPDuration - Int(MPDuration) - 0.5) > 0 Then
        SECONDS_LEFT = Left(Int(MPDuration) + 1, 2)
    Else
        SECONDS_LEFT = Left(Int(MPDuration), 2)
    End If
    If Len(SECONDS_LEFT) = 1 Then SECONDS_LEFT = "0" + SECONDS_LEFT
End Sub
'------------------------------------------'
'PREPARE FORM TO ADD DIRECTORY TO PLAYLIST '
'------------------------------------------'
Public Sub PREPARE_ADD_DIR()
' Prepare form to add a directory
    frmMain.List1.Visible = False
    frmMain.Dir1.Visible = True
    frmMain.Drive1.Visible = True
    frmMain.cmdCancel.Visible = True
    frmMain.cmdOK.Visible = True
    frmMain.cmdLup.Visible = True
    frmMain.cmdMyDocuments.Visible = True
    frmMain.cmdDesktop.Visible = True
End Sub
Public Sub RETURN_ADD_DIR()
' Return form from add dir condition
    frmMain.List1.Visible = True
    frmMain.Dir1.Visible = False
    frmMain.Drive1.Visible = False
    frmMain.cmdCancel.Visible = False
    frmMain.cmdOK.Visible = False
    frmMain.cmdLup.Visible = False
    frmMain.cmdMyDocuments.Visible = False
    frmMain.cmdDesktop.Visible = False
End Sub
'----------------------------'
'         SEQUENCER          '
'----------------------------'
Public Sub SEQUENCE()
' Choose next song for graphic display (Next/Random/None) and set equalizer to zero
    frmMain.cmdNextList.Picture = frmMain.ImgNextListoff.Picture
    frmMain.cmdRndList.Picture = frmMain.ImgRndListoff.Picture
    frmMain.cmdNoneList.Picture = frmMain.ImgNoneListoff.Picture
    Select Case NEXT_SONG
    Case Is = 1
        frmMain.cmdNextList.Picture = frmMain.ImgNextListon.Picture
    Case Is = 2
        frmMain.cmdRndList.Picture = frmMain.ImgRndListon.Picture
    Case Is = 3
        frmMain.cmdNoneList.Picture = frmMain.ImgNoneListon.Picture
    End Select
End Sub
'----------------------------'
'   ARRANGE SONG TITLE BAR   '
'----------------------------'
Public Sub FIX_TITLE_BAR()
' Fix title bar size
    If Len(frmMain.TxtName.Text) > 33 Then
        frmMain.TxtName.Text = "..." & Right(frmMain.TxtName.Text, 30)
    End If
End Sub
'----------------------------'
' AUTOSAVE PLAYLIST ON EXIT  '
'----------------------------'
Public Sub AUTOSAVE_PLAYLIST()
' Save playlist on exit to auto load it later
On Error GoTo Problem
     Open (AUTOLIST) For Output As #1
     Print #1, "Rasputin Playlist : AutoList"
     Dim I%
     For I = 0 To frmMain.List1.ListCount - 1
     Print #1, frmMain.List1.List(I)
     Next
     Close #1
     Exit Sub
Problem:
MsgBox "An error appeared while trying to save" & vbNewLine & "the auto list and the file was not created.", vbOKOnly & vbInformation, "Error while saving auto playlist"
Exit Sub
End Sub
'----------------------------'
'AUTOLOAD PLAYLIST ON STARTUP'
'----------------------------'
Public Sub AUTOLOAD_PLAYLIST()
' Load playlist used in the program's last use
    If AUTOLOAD_LIST = True Then
        On Error GoTo Problem
        Open AUTOLIST For Input As #1
        Input #1, SONG$
        If Left(SONG$, 20) <> "Rasputin Playlist : " Then GoTo Problem
        Do Until EOF(1)
            Input #1, SONG$
            frmMain.List1.AddItem SONG$
        Loop
        Close 1
        frmMain.List1.ListIndex = frmMain.List1.ListCount - 1
        frmMain.List1_Click
        Exit Sub
Problem:
        MsgBox "The auto list appears to be corrupted or some files" & vbNewLine & "are missing from their original location.", vbOKOnly & vbInformation, "Error while loading auto playlist"
        Exit Sub
    End If
End Sub
'----------------------------'
'    INITIALIZE EQUALIZER    '
'----------------------------'
Public Sub INITIALIZE_EQUALIZER()
' Initialize equalizer bars
    frmMain.Equ.Enabled = True
    For I = 0 To 4 Step 1
        frmMain.Preamp(I).Enabled = True
    Next I
    If vol.VolBassMax <= vol.VolTrebleMax Then
        If vol.VolBassMax >= 32767 Then
            frmMain.Equ.Max = 32767
        Else
            frmMain.Equ.Max = vol.VolBassMax
        End If
    Else
        If vol.VolTrebleMax >= 32767 Then
            frmMain.Equ.Max = 32767
        Else
            frmMain.Equ.Max = vol.VolTrebleMax
        End If
    End If
    frmMain.Equ.Value = frmMain.Equ.Max
End Sub
'----------------------------------'
'LOAD STARTUP OPTIONS FROM REGISTRY'
'----------------------------------'
Public Sub LOAD_OPTIONS()
' Load Options file
On Error GoTo Problem
    ' RW and FF values
    RW = Val(GetSetting(App.Title, "Startup", "RW", DEFAULT_RW))
    FF = Val(GetSetting(App.Title, "Startup", "FF", DEFAULT_FF))
    ' Next song
    NEXT_SONG = Val(GetSetting(App.Title, "Startup", "NEXT_SONG", DEFAULT_NEXT_SONG))
    If NEXT_SONG <> 1 And NEXT_SONG <> 2 And NEXT_SONG <> 3 Then NEXT_SONG = DEFAULT_NEXT_SONG
    Select Case NEXT_SONG
    Case Is = 1
        frmMain.mnuSequenceNextSong.Checked = True
        frmMain.mnuSequenceRnd.Checked = False
        frmMain.mnuSequenceNone.Checked = False
    Case Is = 2
        frmMain.mnuSequenceNextSong.Checked = False
        frmMain.mnuSequenceNone.Checked = False
        frmMain.mnuSequenceRnd.Checked = True
    Case Is = 3
        frmMain.mnuSequenceNextSong.Checked = False
        frmMain.mnuSequenceRnd.Checked = False
        frmMain.mnuSequenceNone.Checked = True
    End Select
    ' Autolist filename
    AUTOLIST = GetSetting(App.Title, "Startup", "AUTO_LIST_FILE", App.Path & "\Autolist.rpl")
    TIMER_SPEED = Val(GetSetting(App.Title, "Startup", "TIMER_SPEED", DEFAULT_TIMER_SPEED))
    ' Autoload list
    Select Case Val(GetSetting(App.Title, "Startup", "AUTO_LIST", DEFAULT_AUTOLOAD_LIST))
    Case Is = 0
        AUTOLOAD_LIST = False
    Case Is = 1
        AUTOLOAD_LIST = True
    Case Else
        AUTOLOAD_LIST = True
    End Select
    ' Load tools
    Select Case Val(GetSetting(App.Title, "Startup", "TOOLS", DEFAULT_FORM_TOOLS))
    Case Is = 0
        FORM_TOOLS = False
    Case Is = 1
        FORM_TOOLS = True
    Case Else
        FORM_TOOLS = False
    End Select
    ' Compact mode
    Select Case Val(GetSetting(App.Title, "Startup", "COMPACT", DEFAULT_FORM_MODE))
    Case Is = 0
        FORM_MODE = False
    Case Is = 1
        FORM_MODE = True
    Case Else
        FORM_MODE = False
    End Select
    ' Form on top
    Select Case Val(GetSetting(App.Title, "Startup", "ON_TOP", DEFAULT_FORM_ON_TOP))
    Case Is = 0
        FORM_ON_TOP = False
    Case Is = 1
        FORM_ON_TOP = True
    Case Else
        FORM_ON_TOP = False
    End Select
    ' Scroll title
    Select Case Val(GetSetting(App.Title, "Startup", "TITLE_SCROLL", DEFAULT_TITLE_SCROLL))
    Case Is = 0
        TITLE_SCROLL = False
    Case Is = 1
        TITLE_SCROLL = True
    Case Else
        TITLE_SCROLL = False
    End Select
    ' Minimize mode
    Select Case Val(GetSetting(App.Title, "Startup", "MINIMIZE_MODE", DEFAULT_MINIMIZE_MODE))
    Case Is = 0
        MINIMIZE_MODE = False
    Case Is = 1
        MINIMIZE_MODE = True
    Case Else
        MINIMIZE_MODE = False
    End Select
    ' Autoplay
    Select Case Val(GetSetting(App.Title, "Startup", "AUTOPLAY", DEFAULT_AUTOPLAY))
    Case Is = 0
        AUTOPLAY = False
    Case Is = 1
        AUTOPLAY = True
    Case Else
        AUTOPLAY = DEFAULT_AUTOPLAY
    End Select
    Exit Sub
Problem:
    ' On error load default values
    MsgBox 1
    AUTOLOAD_LIST = DEFAULT_AUTOLOAD_LIST
    AUTOLIST = App.Path & "\Autolist.rpl"
    FORM_ON_TOP = DEFAULT_FORM_ON_TOP
    TITLE_SCROLL = DEFAULT_TITLE_SCROLL
    FORM_MODE = DEFAULT_FORM_MODE
    FORM_TOOLS = DEFAULT_FORM_TOOLS
    NEXT_SONG = DEFAULT_NEXT_SONG
    Exit Sub
End Sub
'-----------------------------'
'PREPARE MAIN FORM FOR CHANGES'
'-----------------------------'
Public Sub PREPARE_MAIN_FORM()
    ' Set equalizer to zero
    frmMain.ListPresets.Visible = False
    EquTemp = 32767
    For I = 0 To 4 Step 1
        Equ5temp(I) = 0
        frmMain.Preamp(I).Value = 0
    Next I
    ' Form on top check
    If FORM_ON_TOP = True Then
        FORM_ON_TOP = False
        frmMain.CmdOnTop_Click
    End If
    If FORM_ON_TOP = True Then
        frmMain.mnuOnTop.Checked = True
    Else
        frmMain.mnuOnTop.Checked = False
    End If
    ' Form mode check
    Select Case FORM_MODE
    Case Is = True
        ' Compact mode
        FORM_MODE = False
        frmMain.cmdCompact_Click
    Case Is = False
        ' Tools/List mode
        If FORM_TOOLS = True Then
            FORM_TOOLS = False
            frmMain.cmdTools_Click
        End If
    End Select
End Sub
'---------------------'
'SET EQUALIZER PRESETS'
'---------------------'
Public Sub SET_PRESET(preset_type As Integer)
Select Case preset_type
Case Is = 0
    ' [Current] Do nothing and return
Case Is = 1
    ' Blues
    frmMain.Preamp(0).Value = 20
    frmMain.Preamp(1).Value = 5
    frmMain.Preamp(2).Value = 10
    frmMain.Preamp(3).Value = 20
    frmMain.Preamp(4).Value = 30
Case Is = 2
    ' Classic
    frmMain.Preamp(0).Value = 0
    frmMain.Preamp(1).Value = -5
    frmMain.Preamp(2).Value = -10
    frmMain.Preamp(3).Value = -20
    frmMain.Preamp(4).Value = -35
Case Is = 3
    ' Disco
    frmMain.Preamp(0).Value = 10
    frmMain.Preamp(1).Value = 0
    frmMain.Preamp(2).Value = -10
    frmMain.Preamp(3).Value = -5
    frmMain.Preamp(4).Value = 5
Case Is = 4
    ' Greek Folk
    frmMain.Preamp(0).Value = 10
    frmMain.Preamp(1).Value = 5
    frmMain.Preamp(2).Value = 0
    frmMain.Preamp(3).Value = -5
    frmMain.Preamp(4).Value = 10
Case Is = 5
    ' Hall
    frmMain.Preamp(0).Value = -30
    frmMain.Preamp(1).Value = -15
    frmMain.Preamp(2).Value = 0
    frmMain.Preamp(3).Value = -5
    frmMain.Preamp(4).Value = 0
Case Is = 6
    ' Hip Hop / Rap
    frmMain.Preamp(0).Value = 5
    frmMain.Preamp(1).Value = 0
    frmMain.Preamp(2).Value = 5
    frmMain.Preamp(3).Value = 15
    frmMain.Preamp(4).Value = 30
Case Is = 7
    ' Heavy Metal
    frmMain.Preamp(0).Value = 30
    frmMain.Preamp(1).Value = 25
    frmMain.Preamp(2).Value = 30
    frmMain.Preamp(3).Value = 40
    frmMain.Preamp(4).Value = 50
Case Is = 8
    ' Jazz
    frmMain.Preamp(0).Value = -10
    frmMain.Preamp(1).Value = -5
    frmMain.Preamp(2).Value = -10
    frmMain.Preamp(3).Value = 0
    frmMain.Preamp(4).Value = 10
Case Is = 9
    ' Live
    frmMain.Preamp(0).Value = -15
    frmMain.Preamp(1).Value = 0
    frmMain.Preamp(2).Value = 15
    frmMain.Preamp(3).Value = 5
    frmMain.Preamp(4).Value = 0
Case Is = 10
    ' Movie
    frmMain.Preamp(0).Value = 20
    frmMain.Preamp(1).Value = 20
    frmMain.Preamp(2).Value = 20
    frmMain.Preamp(3).Value = 10
    frmMain.Preamp(4).Value = 0
Case Is = 11
    ' Pop
    frmMain.Preamp(0).Value = 5
    frmMain.Preamp(1).Value = 10
    frmMain.Preamp(2).Value = 15
    frmMain.Preamp(3).Value = 20
    frmMain.Preamp(4).Value = 0
Case Is = 12
    ' Rave
    frmMain.Preamp(0).Value = 10
    frmMain.Preamp(1).Value = -10
    frmMain.Preamp(2).Value = 10
    frmMain.Preamp(3).Value = 15
    frmMain.Preamp(4).Value = 20
Case Is = 13
    ' Rock
    frmMain.Preamp(0).Value = 30
    frmMain.Preamp(1).Value = 15
    frmMain.Preamp(2).Value = 25
    frmMain.Preamp(3).Value = 30
    frmMain.Preamp(4).Value = 35
End Select
End Sub
'--------------------'
' PAUSE/WAIT SECONDS '
'--------------------'
Public Sub WAIT(WaitingTime)
' Make program wait for WaitingTime seconds
Dim StartWaiting, FinishWaiting
   StartWaiting = Timer
   Do While Timer < StartWaiting + WaitingTime
      DoEvents
   Loop
   FinishWaiting = Timer
End Sub
'---------------------------------'
'DETECT SOUND CARD OR EXIT PROGRAM'
'---------------------------------'
Public Sub DETECT_SOUND_CARD()
' Check for the presence of a sound card (if there isn't one, exit)
    Dim TestVariable As Long
    TestVariable = waveOutGetNumDevs()
    If TestVariable <= 0 Then
        RESPONCE = MsgBox("Your system don't have a sound card and" & vbNewLine & "it cannot play audio files.Rasputin Player is exitting...", vbOKOnly & vbInformation, "Sound Card Error")
        End
    End If
End Sub
'---------------------------'
'GRAPHIC EQUALIZER FUNCTIONS'
'---------------------------'
Public Sub DoReverse()
' Reverse array
    Dim I As Long
    For I = LBound(ReversedBits) To UBound(ReversedBits)
        ReversedBits(I) = ReverseBits(I, NUMBITS)
    Next
End Sub
Public Function ReverseBits(ByVal Index As Long, NUMBITS As Byte) As Long
' Reverse Bits
    Dim I As Byte, Rev As Long
    For I = 0 To NUMBITS - 1
        Rev = (Rev * 2) Or (Index And 1)
        Index = Index \ 2
    Next
    ReverseBits = Rev
End Function
Public Sub FFTAudio(RealIn() As Integer, RealOut() As Single)
' Audio Handling
    Static ImagOut(0 To NUMSAMPLES - 1) As Single
    Static I As Long, j As Long, k As Long, n As Long, BlockSize As Long, BlockEnd As Long
    Static DeltaAngle As Single, DeltaAr As Single
    Static Alpha As Single, Beta As Single
    Static TR As Single, TI As Single, AR As Single, AI As Single
    For I = 0 To (NUMSAMPLES - 1)
        j = ReversedBits(I)
        RealOut(j) = RealIn(I)
        ImagOut(j) = 0
    Next
    BlockEnd = 1
    BlockSize = 2
    Do While BlockSize <= NUMSAMPLES
        DeltaAngle = ANGLENUMERATOR / BlockSize
        Alpha = Sin(0.5 * DeltaAngle)
        Alpha = 2! * Alpha * Alpha
        Beta = Sin(DeltaAngle)
        I = 0
        Do While I < NUMSAMPLES
            AR = 1!
            AI = 0!
            j = I
            For n = 0 To BlockEnd - 1
                k = j + BlockEnd
                TR = AR * RealOut(k) - AI * ImagOut(k)
                TI = AI * RealOut(k) + AR * ImagOut(k)
                RealOut(k) = RealOut(j) - TR
                ImagOut(k) = ImagOut(j) - TI
                RealOut(j) = RealOut(j) + TR
                ImagOut(j) = ImagOut(j) + TI
                DeltaAr = Alpha * AR + Beta * AI
                AI = AI - (Alpha * AI - Beta * AR)
                AR = AR - DeltaAr
                j = j + 1
            Next n
            I = I + BlockSize
        Loop
        BlockEnd = BlockSize
        BlockSize = BlockSize * 2
    Loop
End Sub
'------------------'
'FILE INFO FUNTIONS'
'------------------'
Public Function SHIFT_BITS(Din As String) As String
' Left shift 4 bits losing most significant 4 bits
     Dim SHIFT_ONE As Integer    ' Shift bits
     Dim SHIFT_TWO As Integer    ' Shift bits
     Dim SD1 As Integer          ' Left bit
     Dim SD2 As Integer          ' Middle bit
     Dim SD3 As Integer          ' Right bit
     SD1 = Asc(Left(Din, 1))
     SD2 = Asc(Mid(Din, 2, 1))
     SD3 = Asc(Right(Din, 1))
     SHIFT_ONE = ((SD1 And &HF) * 16) Or ((SD2 And &HF0) / 16)
     SHIFT_TWO = ((SD2 And &HF) * 16) Or ((SD3 And &HF0) / 16)
     SHIFT_BITS = Chr(SHIFT_ONE) + Chr(SHIFT_TWO)
End Function
Public Sub PREPARE_FILE_INFO()
' Setup array for mpeg bitrate info
Dim X As Integer    ' Counter
Dim Y As Integer    ' Counter
Dim BITRATE_DATA    ' Bitrate Info Array
    BITRATE_DATA = "032,032,032,032,008,008,064,048,040,"
    BITRATE_DATA = BITRATE_DATA & "048,016,016,096,056,048,"
    BITRATE_DATA = BITRATE_DATA & "056,024,024,128,064,056,"
    BITRATE_DATA = BITRATE_DATA & "064,032,032,160,080,064,"
    BITRATE_DATA = BITRATE_DATA & "080,040,040,192,096,080,"
    BITRATE_DATA = BITRATE_DATA & "096,048,048,224,112,096,"
    BITRATE_DATA = BITRATE_DATA & "112,056,056,256,128,112,"
    BITRATE_DATA = BITRATE_DATA & "128,064,064,288,160,128,"
    BITRATE_DATA = BITRATE_DATA & "144,080,080,320,192,160,"
    BITRATE_DATA = BITRATE_DATA & "160,096,096,352,224,192,"
    BITRATE_DATA = BITRATE_DATA & "176,112,112,384,256,224,"
    BITRATE_DATA = BITRATE_DATA & "192,128,128,416,320,256,"
    BITRATE_DATA = BITRATE_DATA & "224,144,144,448,384,320,"
    BITRATE_DATA = BITRATE_DATA & "256,160,160,"
    For Y = 1 To 14 Step 1
        For X = 7 To 5 Step -1
            BITRATE_LOOKUP(X, Y) = Left(BITRATE_DATA, 3)
            BITRATE_DATA = Right(BITRATE_DATA, Len(BITRATE_DATA) - 4)
        Next X
        For X = 3 To 1 Step -1
            BITRATE_LOOKUP(X, Y) = Left(BITRATE_DATA, 3)
            BITRATE_DATA = Right(BITRATE_DATA, Len(BITRATE_DATA) - 4)
        Next X
    Next Y
End Sub
Public Sub Getmp3data(MP3File As String)
    Dim Din                       ' Reading string
    Dim Byte1 As Integer          ' Byte read
    Dim Byte2 As Integer          ' Byte read
    Dim Byte3 As Integer          ' Byte read
    Dim Byte4 As Integer          ' Byte read
    Dim Mp3_ID                    ' MPEG significant bit - 1st bit is ID : 0=MPG-2, 1=MPG-1
    Dim Mp3_LAYER                 ' MPEG significant bit - Next 2 bits are Layer
    Dim Mp3_PROT                  ' MPEG significant bit - Next 1 bit is Protection
    Dim Mp3_BITRATE               ' MPEG significant bit - Next 4 bits are bitrate
    Dim Mp3_FREQ                  ' MPEG significant bit - Next 2 bits are frequency
    Dim Mp3_PAD                   ' MPEG significant bit - Next 1 bit is Padding
    Dim Temp_string As String     ' Temporary string variable
    Dim Mp3bits_string As String  ' Temporary Bitrate string variable
    Dim Sample_Rate As Long       ' Sample rate
    Dim DShift As String          ' Shifted bytes
    Dim FSize                     ' File size
    On Error GoTo Problem
    PREPARE_FILE_INFO
    Open MP3File For Binary As #1
        Din = Input(4096, #1)   ' Read 4K
        FSize = LOF(1)
        frmMain.lbl_FI_SZ.Caption = Round(LOF(1) / 1048576, 3) & " Mb"
    Close #1
    I = 0
    Do Until I = 4095
        I = I + 1
        Byte1 = Asc(Mid(Din, I, 1))
        Byte2 = Asc(Mid(Din, I + 1, 1))
        If Byte1 = &HFF And (Byte2 And &HF0) = &HF0 Then
            Temp_string = Mid(Din, I + 1, 3)
            Mp3bits_string = SHIFT_BITS(Mid(Din, I + 1, 3))
            Exit Do
        End If
        DShift = SHIFT_BITS(Mid(Din, I, 3))
        Byte3 = Asc(Left(DShift, 1))
        Byte4 = Asc(Right(DShift, 1))
        If Byte3 = &HFF And (Byte4 And &HF0) = &HF0 Then
            Mp3bits_string = Mid(Din, I + 2, 3)
            Exit Do
        End If
    Loop
    Mp3_ID = (&H80 And Asc(Left(Mp3bits_string, 1))) / 128
    Mp3_LAYER = (&H60 And Asc(Left(Mp3bits_string, 1))) / 32
    Mp3_PROT = &H10 And Asc(Left(Mp3bits_string, 1))
    Mp3_BITRATE = &HF And Asc(Left(Mp3bits_string, 1))
    Mp3_FREQ = &HC0 And Asc(Mid(Mp3bits_string, 2, 1))
    Mp3_PAD = (&H20 And Asc(Mid(Mp3bits_string, 2, 1))) / 2
    ACTUAL_BITRATE = 1000 * CLng((BITRATE_LOOKUP((Mp3_ID * 4) Or Mp3_LAYER, Mp3_BITRATE)))
    If Mp3_ID = 0 Then
        frmMain.lbl_FI_ID.Caption = "MPEG-2"
    Else
        frmMain.lbl_FI_ID.Caption = "MPEG-1"
    End If
    Select Case Mp3_LAYER
        Case 1
            frmMain.lbl_FI_ID.Caption = frmMain.lbl_FI_ID.Caption & " Layer III"
        Case 2
            frmMain.lbl_FI_ID.Caption = frmMain.lbl_FI_ID.Caption & " Layer II"
        Case 3
            frmMain.lbl_FI_ID.Caption = frmMain.lbl_FI_ID.Caption & " Layer I"
    End Select
    Select Case (Mp3_ID * 4) Or Mp3_FREQ
        Case 0
            Sample_Rate = 22050
        Case 1
            Sample_Rate = 24000
        Case 2
            Sample_Rate = 16000
        Case 4
            Sample_Rate = 44100
        Case 5
            Sample_Rate = 48000
        Case 6
            Sample_Rate = 32000
    End Select
    frmMain.lbl_FI_FR.Caption = FSize / (((144 * ACTUAL_BITRATE) / Sample_Rate) + Mp3_PAD) & " Frames"
    frmMain.lbl_FI_BR.Caption = Str(ACTUAL_BITRATE / 1000) & " Kb"
    frmMain.lbl_FI_SR.Caption = Str(Sample_Rate / 1000) & " Khz"
    Exit Sub
Problem:
    frmMain.lbl_FI_ID.Caption = "Unknown"
    frmMain.lbl_FI_BR.Caption = "Unknown"
    frmMain.lbl_FI_SR.Caption = "Unknown"
End Sub
'---------------------------'
'GRAPHIC EQUALIZER FUNCTIONS'
'---------------------------'
Public Sub DoStop()
' Stop Greq
    Call waveInReset(DevHandle)
    Call waveInClose(DevHandle)
    DevHandle = 0
End Sub
