VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVolume"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'----------------'
'  VOLUME CLASS  '
'----------------'
' Constants
Private Const MMSYSERR_NOERROR = 0
Private Const MAXPNAMELEN = 32
Private Const MIXER_LONG_NAME_CHARS = 64
Private Const MIXER_SHORT_NAME_CHARS = 16
Private Const MIXER_GETLINEINFOF_COMPONENTTYPE = &H3&
Private Const MIXER_GETCONTROLDETAILSF_VALUE = &H0&
Private Const MIXER_GETLINECONTROLSF_ONEBYTYPE = &H2&
Private Const MIXER_GETLINECONTROLSF_ALL = &H0&
Private Const MIXER_GETLINECONTROLSF_ONEBYID = &H1&
Private Const MIXER_GETLINECONTROLSF_QUERYMASK = &HF&
Private Const MIXERLINE_COMPONENTTYPE_DST_FIRST = &H0&
Private Const MIXERLINE_COMPONENTTYPE_SRC_FIRST = &H1000&
Private Const MIXER_SETCONTROLDETAILSF_VALUE = &H0&
Private Const MIXERCONTROL_CT_CLASS_FADER = &H50000000
Private Const MIXERCONTROL_CT_UNITS_UNSIGNED = &H30000
Private Const MIXERCONTROL_CT_CLASS_LIST = &H70000000
Private Const MIXERCONTROL_CT_SC_LIST_MULTIPLE = &H1000000
Private Const MIXERCONTROL_CT_UNITS_BOOLEAN = &H10000
Private Const MIXERCONTROL_CT_CLASS_SWITCH = &H20000000
Private Const MIXERCONTROL_CT_SC_SWITCH_BOOLEAN = &H0&
Private Const MIXERCONTROL_CT_CLASS_METER = &H10000000
Private Const MIXERCONTROL_CT_SC_METER_POLLED = &H0&
Private Const MIXERCONTROL_CT_UNITS_SIGNED = &H20000
Private Const MIXERCONTROL_CT_CLASS_SLIDER = &H40000000
Private Const MIXERLINE_COMPONENTTYPE_DST_SPEAKERS = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 4)
Private Const MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 3)
Private Const MIXERLINE_COMPONENTTYPE_SRC_LINE = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 2)
Private Const MIXERLINE_COMPONENTTYPE_SRC_COMPACTDISC = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 5)
Private Const MIXERLINE_COMPONENTTYPE_SRC_SYNTHESIZER = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 4)
Private Const MIXERLINE_COMPONENTTYPE_DST_WAVEIN = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 7)
Private Const MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT = (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 8)
Private Const MIXERCONTROL_CONTROLTYPE_SIGNEDMETER = (MIXERCONTROL_CT_CLASS_METER Or MIXERCONTROL_CT_SC_METER_POLLED Or MIXERCONTROL_CT_UNITS_SIGNED)
Private Const MIXERCONTROL_CONTROLTYPE_BOOLEAN = (MIXERCONTROL_CT_CLASS_SWITCH Or MIXERCONTROL_CT_SC_SWITCH_BOOLEAN Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Private Const MIXERCONTROL_CONTROLTYPE_MULTIPLESELECT = (MIXERCONTROL_CT_CLASS_LIST Or MIXERCONTROL_CT_SC_LIST_MULTIPLE Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Private Const MIXERCONTROL_CONTROLTYPE_FADER = (MIXERCONTROL_CT_CLASS_FADER Or MIXERCONTROL_CT_UNITS_UNSIGNED)
Private Const MIXERCONTROL_CONTROLTYPE_VOLUME = (MIXERCONTROL_CONTROLTYPE_FADER + 1)
Private Const MIXERCONTROL_CONTROLTYPE_BASS = (MIXERCONTROL_CONTROLTYPE_FADER + 2)
Private Const MIXERCONTROL_CONTROLTYPE_EQUALIZER = (MIXERCONTROL_CONTROLTYPE_FADER + 4)
Private Const MIXERCONTROL_CONTROLTYPE_MIXER = (MIXERCONTROL_CONTROLTYPE_MULTIPLESELECT + 1)
Private Const MIXERCONTROL_CONTROLTYPE_MONO = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 3)
Private Const MIXERCONTROL_CONTROLTYPE_MUTE = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 2)
Private Const MIXERCONTROL_CONTROLTYPE_ONOFF = (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 1)
Private Const MIXERCONTROL_CONTROLTYPE_PEAKMETER = (MIXERCONTROL_CONTROLTYPE_SIGNEDMETER + 1)
Private Const MIXERCONTROL_CONTROLTYPE_TREBLE = (MIXERCONTROL_CONTROLTYPE_FADER + 3)
Private Const MIXERCONTROL_CONTROLTYPE_SLIDER = (MIXERCONTROL_CT_CLASS_SLIDER Or MIXERCONTROL_CT_UNITS_SIGNED)
Private Const MIXERCONTROL_CONTROLTYPE_PAN = (MIXERCONTROL_CONTROLTYPE_SLIDER + 1)
Private Const CALLBACK_FUNCTION = &H30000
Private Const MM_WIM_DATA = &H3C0
Private Const WHDR_DONE = &H1
Private Const GMEM_FIXED = &H0
Private Const NUM_BUFFERS = 2
Private Const BUFFER_SIZE = 8192
Private Const DEVICEID = 0
' Variables
Private hmixer As Long
Private volCtrl As MIXERCONTROL
Private wavCtrl As MIXERCONTROL
Private volMute As MIXERCONTROL
Private wavMute As MIXERCONTROL
Private volFad As MIXERCONTROL
Private wavFad As MIXERCONTROL
Private volBass As MIXERCONTROL
Private wavBass As MIXERCONTROL
Private volTreb As MIXERCONTROL
Private wavTreb As MIXERCONTROL
Private volPan As MIXERCONTROL
Private wavPan As MIXERCONTROL
Private rc As Long
Private ok As Boolean
Private I As Integer
Private j As Integer
Private msg As String * 200
Private hWaveIn As Long
Private format As WAVEFORMAT
Private hmem(NUM_BUFFERS) As Long
Private inHdr(NUM_BUFFERS) As WAVEHDR
Private fRecording As Boolean
Private inputVolCtrl As MIXERCONTROL
Private outputVolCtrl As MIXERCONTROL
Private vol As MIXERCONTROLDETAILS_SIGNED
Private volume As Long
Private volHmem As Long
Private mxcd As MIXERCONTROLDETAILS
' Declarations
Private Declare Function mixerClose Lib "winmm.dll" (ByVal hmx As Long) As Long
Private Declare Function mixerGetControlDetails Lib "winmm.dll" Alias "mixerGetControlDetailsA" (ByVal hmxobj As Long, pMxcd As MIXERCONTROLDETAILS, ByVal fdwDetails As Long) As Long
Private Declare Function mixerGetDevCaps Lib "winmm.dll" Alias "mixerGetDevCapsA" (ByVal uMxId As Long, ByVal pmxcaps As MIXERCAPS, ByVal cbmxcaps As Long) As Long
Private Declare Function mixerGetID Lib "winmm.dll" (ByVal hmxobj As Long, pumxID As Long, ByVal fdwId As Long) As Long
Private Declare Function mixerGetLineControls Lib "winmm.dll" Alias "mixerGetLineControlsA" (ByVal hmxobj As Long, pmxlc As MIXERLINECONTROLS, ByVal fdwControls As Long) As Long
Private Declare Function mixerGetLineInfo Lib "winmm.dll" Alias "mixerGetLineInfoA" (ByVal hmxobj As Long, pmxl As MIXERLINE, ByVal fdwInfo As Long) As Long
Private Declare Function mixerGetNumDevs Lib "winmm.dll" () As Long
Private Declare Function mixerMessage Lib "winmm.dll" (ByVal hmx As Long, ByVal uMsg As Long, ByVal dwParam1 As Long, ByVal dwParam2 As Long) As Long
Private Declare Function mixerOpen Lib "winmm.dll" (phmx As Long, ByVal uMxId As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal fdwOpen As Long) As Long
Private Declare Function mixerSetControlDetails Lib "winmm.dll" (ByVal hmxobj As Long, pMxcd As MIXERCONTROLDETAILS, ByVal fdwDetails As Long) As Long
Private Declare Sub CopyStructFromPtr Lib "kernel32" Alias "RtlMoveMemory" (struct As Any, ByVal ptr As Long, ByVal cb As Long)
Private Declare Sub CopyPtrFromStruct Lib "kernel32" Alias "RtlMoveMemory" (ByVal ptr As Long, struct As Any, ByVal cb As Long)
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hmem As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hmem As Long) As Long
Private Declare Function waveInOpen Lib "winmm.dll" (lphWaveIn As Long, ByVal uDeviceID As Long, lpFormat As WAVEFORMAT, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Private Declare Function waveInPrepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Private Declare Function waveInReset Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Private Declare Function waveInStart Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Private Declare Function waveInStop Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Private Declare Function waveInUnprepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Private Declare Function waveInClose Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Private Declare Function waveInGetDevCaps Lib "winmm.dll" Alias "waveInGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As WAVEINCAPS, ByVal uSize As Long) As Long
Private Declare Function waveInGetNumDevs Lib "winmm.dll" () As Long
Private Declare Function waveInGetErrorText Lib "winmm.dll" Alias "waveInGetErrorTextA" (ByVal err As Long, ByVal lpText As String, ByVal uSize As Long) As Long
Private Declare Function waveInAddBuffer Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
' Types
Private Type MIXERCAPS
    wMid As Integer
    wPid As Integer
    vDriverVersion As Long
    szPname As String * MAXPNAMELEN
    fdwSupport As Long
    cDestinations As Long
End Type
Private Type MIXERCONTROL
    cbStruct As Long
    dwControlID As Long
    dwControlType As Long
    fdwControl As Long
    cMultipleItems As Long
    szShortName As String * MIXER_SHORT_NAME_CHARS
    szName As String * MIXER_LONG_NAME_CHARS
    lMinimum As Long
    lMaximum As Long
    Reserved(10) As Long
End Type
Private Type MIXERCONTROLDETAILS
    cbStruct As Long
    dwControlID As Long
    cChannels As Long
    item As Long
    cbDetails As Long
    paDetails As Long
End Type
Private Type MIXERCONTROLDETAILS_UNSIGNED
    dwValue As Long
End Type
Private Type MIXERCONTROLDETAILS_BOOLEAN
    fValue As Long
End Type
Private Type MIXERCONTROLDETAILS_SIGNED
   lValue As Long
End Type
Private Type MIXERCONTROLDETAILS_LISTTEXT
    dwParam1 As Long
    dwParam2 As Long
    szName As String * MIXER_LONG_NAME_CHARS
End Type
Private Type MIXERLINE
    cbStruct As Long
    dwDestination As Long
    dwSource As Long
    dwLineID As Long
    fdwLine As Long
    dwUser As Long
    dwComponentType As Long
    cChannels As Long
    cConnections As Long
    cControls As Long
    szShortName As String * MIXER_SHORT_NAME_CHARS
    szName As String * MIXER_LONG_NAME_CHARS
    dwType As Long
    dwDeviceID As Long
    wMid  As Integer
    wPid As Integer
    vDriverVersion As Long
    szPname As String * MAXPNAMELEN
End Type
Private Type MIXERLINECONTROLS
    cbStruct As Long
    dwLineID As Long
    dwControl As Long
    cControls As Long
    cbmxctrl As Long
    pamxctrl As Long
End Type
Private Type WAVEHDR
   lpData As Long
   dwBufferLength As Long
   dwBytesRecorded As Long
   dwUser As Long
   dwFlags As Long
   dwLoops As Long
   lpNext As Long
   Reserved As Long
End Type
Private Type WAVEINCAPS
   wMid As Integer
   wPid As Integer
   vDriverVersion As Long
   szPname As String * 32
   dwFormats As Long
   wChannels As Integer
End Type
Private Type WAVEFORMAT
   wFormatTag As Integer
   nChannels As Integer
   nSamplesPerSec As Long
   nAvgBytesPerSec As Long
   nBlockAlign As Integer
   wBitsPerSample As Integer
   cbSize As Integer
End Type
Private Sub Class_Initialize()
' Apply and initialize controls
    rc = mixerOpen(hmixer, 0, 0, 0, 0)
    ok = GetVolumeControl(hmixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, MIXERCONTROL_CONTROLTYPE_VOLUME, volCtrl)
    ok = GetVolumeControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT, MIXERCONTROL_CONTROLTYPE_VOLUME, wavCtrl)
    ok = GetVolumeControl(hmixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, MIXERCONTROL_CONTROLTYPE_MUTE, volMute)
    ok = GetVolumeControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT, MIXERCONTROL_CONTROLTYPE_MUTE, wavMute)
    ok = GetVolumeControl(hmixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, MIXERCONTROL_CONTROLTYPE_FADER, volFad)
    ok = GetVolumeControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT, MIXERCONTROL_CONTROLTYPE_FADER, wavFad)
    ok = GetVolumeControl(hmixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, MIXERCONTROL_CONTROLTYPE_BASS, volBass)
    ok = GetVolumeControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT, MIXERCONTROL_CONTROLTYPE_BASS, wavBass)
    ok = GetVolumeControl(hmixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, MIXERCONTROL_CONTROLTYPE_TREBLE, volTreb)
    ok = GetVolumeControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT, MIXERCONTROL_CONTROLTYPE_TREBLE, wavTreb)
    ok = GetVolumeControl(hmixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, MIXERCONTROL_CONTROLTYPE_PAN, volPan)
    ok = GetVolumeControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT, MIXERCONTROL_CONTROLTYPE_PAN, wavPan)
   mxcd.cbStruct = Len(mxcd)
   volHmem = GlobalAlloc(&H0, Len(volume))
   mxcd.paDetails = GlobalLock(volHmem)
   mxcd.cbDetails = Len(volume)
   mxcd.cChannels = 1
   ok = GetVolumeControl(hmixer, MIXERLINE_COMPONENTTYPE_DST_WAVEIN, MIXERCONTROL_CONTROLTYPE_PEAKMETER, inputVolCtrl)
   If (ok <> True) Then
       ok = GetVolumeControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE, MIXERCONTROL_CONTROLTYPE_PEAKMETER, inputVolCtrl)
   End If
   ok = GetVolumeControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT, MIXERCONTROL_CONTROLTYPE_PEAKMETER, outputVolCtrl)
End Sub
'-------------------------'
'GET/SET MAIN VOLUME LEVEL'
'-------------------------'
Private Function SetVolume(mxc As MIXERCONTROL, ByVal volume As Long) As Boolean
' Set volume
    Dim mxcd As MIXERCONTROLDETAILS
    Dim vol As MIXERCONTROLDETAILS_UNSIGNED
    Dim hmem As Long
    mxcd.item = 0
    mxcd.dwControlID = mxc.dwControlID
    mxcd.cbStruct = Len(mxcd)
    mxcd.cbDetails = Len(vol)
    hmem = GlobalAlloc(&H40, Len(vol))
    mxcd.paDetails = GlobalLock(hmem)
    mxcd.cChannels = 1
    vol.dwValue = volume
    CopyPtrFromStruct mxcd.paDetails, vol, Len(vol)
    rc = mixerSetControlDetails(hmixer, mxcd, MIXER_SETCONTROLDETAILSF_VALUE)
    GlobalFree (hmem)
    If (MMSYSERR_NOERROR = rc) Then
        SetVolume = True
    Else
        SetVolume = False
    End If
End Function
Private Function GetVolume(mxc As MIXERCONTROL) As Long
' Get volume
    Dim mxcd As MIXERCONTROLDETAILS
    Dim vol As MIXERCONTROLDETAILS_UNSIGNED
    Dim hmem As Long
    mxcd.item = 0
    mxcd.dwControlID = mxc.dwControlID
    mxcd.cbStruct = Len(mxcd)
    mxcd.cbDetails = Len(vol)
    hmem = GlobalAlloc(&H40, Len(vol))
    mxcd.paDetails = GlobalLock(hmem)
    mxcd.cChannels = 1
    rc = mixerGetControlDetails(hmixer, mxcd, MIXER_GETCONTROLDETAILSF_VALUE)
    CopyStructFromPtr vol, mxcd.paDetails, Len(vol)
    GetVolume = vol.dwValue
    GlobalFree (hmem)
End Function
Private Function GetTreble(mxc As MIXERCONTROL) As Long
' Get current treble level
    Dim mxcd As MIXERCONTROLDETAILS
    Dim Treb As MIXERCONTROLDETAILS_UNSIGNED
    Dim hmem As Long
    mxcd.item = 0
    mxcd.dwControlID = mxc.dwControlID
    mxcd.cbStruct = Len(mxcd)
    mxcd.cbDetails = Len(Treb)
    hmem = GlobalAlloc(&H40, Len(Treb))
    mxcd.paDetails = GlobalLock(hmem)
    mxcd.cChannels = 1
    rc = mixerGetControlDetails(hmixer, mxcd, MIXER_GETCONTROLDETAILSF_VALUE)
    CopyStructFromPtr Treb, mxcd.paDetails, Len(Treb)
    GetTreble = Treb.dwValue
    GlobalFree (hmem)
End Function
Private Function SetTreble(mxc As MIXERCONTROL, ByVal Treble As Long)
' Set Treble level
    Dim mxcd As MIXERCONTROLDETAILS
    Dim Treb As MIXERCONTROLDETAILS_UNSIGNED
    Dim hmem As Long
    mxcd.item = 0
    mxcd.dwControlID = mxc.dwControlID
    mxcd.cbStruct = Len(mxcd)
    mxcd.cbDetails = Len(Treb)
    hmem = GlobalAlloc(&H40, Len(Treb))
    mxcd.paDetails = GlobalLock(hmem)
    mxcd.cChannels = 1
    Treb.dwValue = Treble
    CopyPtrFromStruct mxcd.paDetails, Treb, Len(Treb)
    rc = mixerSetControlDetails(hmixer, mxcd, MIXER_SETCONTROLDETAILSF_VALUE)
    GlobalFree (hmem)
End Function
'-------------------------'
'GET/SET BASS VOLUME LEVEL'
'-------------------------'
Private Function GetBass(mxc As MIXERCONTROL) As Long
' Get Bass level
    Dim mxcd As MIXERCONTROLDETAILS
    Dim Bas As MIXERCONTROLDETAILS_UNSIGNED
    Dim hmem As Long
    mxcd.item = 0
    mxcd.dwControlID = mxc.dwControlID
    mxcd.cbStruct = Len(mxcd)
    mxcd.cbDetails = Len(Bas)
    hmem = GlobalAlloc(&H40, Len(Bas))
    mxcd.paDetails = GlobalLock(hmem)
    mxcd.cChannels = 1
    rc = mixerGetControlDetails(hmixer, mxcd, MIXER_GETCONTROLDETAILSF_VALUE)
    CopyStructFromPtr Bas, mxcd.paDetails, Len(Bas)
    GetBass = Bas.dwValue
    GlobalFree (hmem)
End Function
Private Function SetBass(mxc As MIXERCONTROL, ByVal Bass As Long)
' Set bass level
    Dim mxcd As MIXERCONTROLDETAILS
    Dim Bas As MIXERCONTROLDETAILS_UNSIGNED
    Dim hmem As Long
    mxcd.item = 0
    mxcd.dwControlID = mxc.dwControlID
    mxcd.cbStruct = Len(mxcd)
    mxcd.cbDetails = Len(Bas)
    hmem = GlobalAlloc(&H40, Len(Bas))
    mxcd.paDetails = GlobalLock(hmem)
    mxcd.cChannels = 1
    Bas.dwValue = Bass
    CopyPtrFromStruct mxcd.paDetails, Bas, Len(Bas)
    rc = mixerSetControlDetails(hmixer, mxcd, MIXER_SETCONTROLDETAILSF_VALUE)
    GlobalFree (hmem)
End Function
'------------------------'
'GET/SET PAN VOLUME LEVEL'
'------------------------'
Private Function GetPan(mxc As MIXERCONTROL) As Boolean
' Get Pan Level
    Dim mxcd As MIXERCONTROLDETAILS
    Dim Pan As MIXERCONTROLDETAILS_UNSIGNED
    Dim hmem As Long
    mxcd.item = 0
    mxcd.dwControlID = mxc.dwControlID
    mxcd.cbStruct = Len(mxcd)
    mxcd.cbDetails = Len(Pan)
    hmem = GlobalAlloc(&H40, Len(Pan))
    mxcd.paDetails = GlobalLock(hmem)
    mxcd.cChannels = 1
    rc = mixerGetControlDetails(hmixer, mxcd, MIXER_GETCONTROLDETAILSF_VALUE)
    CopyStructFromPtr Pan, mxcd.paDetails, Len(Pan)
    GetPan = Pan.dwValue
    GlobalFree (hmem)
End Function
Private Function SetPan(mxc As MIXERCONTROL, ByVal Pano As Long)
' Set Pan Level
    Dim mxcd As MIXERCONTROLDETAILS
    Dim Pan As MIXERCONTROLDETAILS_UNSIGNED
    Dim hmem As Long
    mxcd.item = 0
    mxcd.dwControlID = mxc.dwControlID
    mxcd.cbStruct = Len(mxcd)
    mxcd.cbDetails = Len(Pan)
    hmem = GlobalAlloc(&H40, Len(Pan))
    mxcd.paDetails = GlobalLock(hmem)
    mxcd.cChannels = 1
    Pan.dwValue = Pano
    CopyPtrFromStruct mxcd.paDetails, Pan, Len(Pan)
    rc = mixerSetControlDetails(hmixer, mxcd, MIXER_SETCONTROLDETAILSF_VALUE)
    GlobalFree (hmem)
End Function
'----------------------'
'GET/SET MUTE CONDITION'
'----------------------'
Private Function GetMute(mxc As MIXERCONTROL) As Boolean
' Get Mute condition
    Dim mxcd As MIXERCONTROLDETAILS
    Dim Mut As MIXERCONTROLDETAILS_BOOLEAN
    Dim hmem As Long
    mxcd.item = 0
    mxcd.dwControlID = mxc.dwControlID
    mxcd.cbStruct = Len(mxcd)
    mxcd.cbDetails = Len(Mut)
    hmem = GlobalAlloc(&H40, Len(Mut))
    mxcd.paDetails = GlobalLock(hmem)
    mxcd.cChannels = 1
    rc = mixerGetControlDetails(hmixer, mxcd, MIXER_GETCONTROLDETAILSF_VALUE)
    CopyStructFromPtr Mut, mxcd.paDetails, Len(Mut)
    GetMute = IIf((Mut.fValue = 1), True, False)
    GlobalFree (hmem)
End Function
Private Function SetMute(mxc As MIXERCONTROL, ByVal Mute As Boolean)
' Set to Mute condition
    Dim mxcd As MIXERCONTROLDETAILS
    Dim Mut As MIXERCONTROLDETAILS_BOOLEAN
    Dim hmem As Long
    mxcd.item = 0
    mxcd.dwControlID = mxc.dwControlID
    mxcd.cbStruct = Len(mxcd)
    mxcd.cbDetails = Len(Mut)
    hmem = GlobalAlloc(&H40, Len(Mut))
    mxcd.paDetails = GlobalLock(hmem)
    mxcd.cChannels = 1
    Mut.fValue = IIf(Mute, 1, 0)
    CopyPtrFromStruct mxcd.paDetails, Mut, Len(Mut)
    rc = mixerSetControlDetails(hmixer, mxcd, MIXER_SETCONTROLDETAILSF_VALUE)
    GlobalFree (hmem)
End Function
'-------------------'
'GET/SET FADER LEVEL'
'-------------------'
Private Function GetFader(mxc As MIXERCONTROL) As Long
' Get Fader level
    Dim mxcd As MIXERCONTROLDETAILS
    Dim Fad As MIXERCONTROLDETAILS_UNSIGNED
    Dim hmem As Long
    mxcd.item = 0
    mxcd.dwControlID = mxc.dwControlID
    mxcd.cbStruct = Len(mxcd)
    mxcd.cbDetails = Len(Fad)
    hmem = GlobalAlloc(&H40, Len(Fad))
    mxcd.paDetails = GlobalLock(hmem)
    mxcd.cChannels = 1
    rc = mixerGetControlDetails(hmixer, mxcd, MIXER_GETCONTROLDETAILSF_VALUE)
    CopyStructFromPtr Fad, mxcd.paDetails, Len(Fad)
    GetFader = Fad.dwValue
    GlobalFree (hmem)
End Function
Private Function SetFader(mxc As MIXERCONTROL, ByVal Fader As Long)
' Set Fader Level
    Dim mxcd As MIXERCONTROLDETAILS
    Dim Fad As MIXERCONTROLDETAILS_UNSIGNED
    Dim hmem As Long
    mxcd.item = 0
    mxcd.dwControlID = mxc.dwControlID
    mxcd.cbStruct = Len(mxcd)
    mxcd.cbDetails = Len(Fad)
    hmem = GlobalAlloc(&H40, Len(Fad))
    mxcd.paDetails = GlobalLock(hmem)
    mxcd.cChannels = 1
    Fad.dwValue = Fader
    CopyPtrFromStruct mxcd.paDetails, Fad, Len(Fad)
    rc = mixerSetControlDetails(hmixer, mxcd, MIXER_SETCONTROLDETAILSF_VALUE)
    GlobalFree (hmem)
End Function
'------------------'
'SOUND MANIPULATION'
'------------------'
Private Function StartInput() As Boolean
' Start inputing / manipulating sound
    If fRecording Then
        StartInput = True
        Exit Function
    End If
    format.wFormatTag = 1
    format.nChannels = 1
    format.wBitsPerSample = 8
    format.nSamplesPerSec = 8000
    format.nBlockAlign = format.nChannels * format.wBitsPerSample / 8
    format.nAvgBytesPerSec = format.nSamplesPerSec * format.nBlockAlign
    format.cbSize = 0
    For I = 0 To NUM_BUFFERS - 1
        hmem(I) = GlobalAlloc(&H40, BUFFER_SIZE)
        inHdr(I).lpData = GlobalLock(hmem(I))
        inHdr(I).dwBufferLength = BUFFER_SIZE
        inHdr(I).dwFlags = 0
        inHdr(I).dwLoops = 0
    Next
    rc = waveInOpen(hWaveIn, DEVICEID, format, 0, 0, 0)
    If rc <> 0 Then
        waveInGetErrorText rc, msg, Len(msg)
        MsgBox msg, vbCritical + vbOKOnly, App.Title & " - Mixer failure"
        StartInput = False
        Exit Function
    End If
    For I = 0 To NUM_BUFFERS - 1
        rc = waveInPrepareHeader(hWaveIn, inHdr(I), Len(inHdr(I)))
        If (rc <> 0) Then
            waveInGetErrorText rc, msg, Len(msg)
            MsgBox msg, vbCritical + vbOKOnly, App.Title & " - Mixer failure"
        End If
    Next
    For I = 0 To NUM_BUFFERS - 1
        rc = waveInAddBuffer(hWaveIn, inHdr(I), Len(inHdr(I)))
        If (rc <> 0) Then
            waveInGetErrorText rc, msg, Len(msg)
            MsgBox msg, vbCritical + vbOKOnly, App.Title & " - Mixer failure"
        End If
    Next
    fRecording = True
    rc = waveInStart(hWaveIn)
    StartInput = True
End Function
Private Sub waveInProc(ByVal hwi As Long, ByVal uMsg As Long, ByVal dwInstance As Long, ByRef hdr As WAVEHDR, ByVal dwParam2 As Long)
' Check if there is a wave in process
   If (uMsg = MM_WIM_DATA) Then
      If fRecording Then
         rc = waveInAddBuffer(hwi, hdr, Len(hdr))
      End If
   End If
End Sub
Private Function GetVolumeControl(ByVal hmixer As Long, ByVal componentType As Long, ByVal ctrlType As Long, ByRef mxc As MIXERCONTROL) As Boolean
' Apply for the volume control
    Dim mxlc As MIXERLINECONTROLS
    Dim mxl As MIXERLINE
    Dim hmem As Long
    Dim rc As Long
    mxl.cbStruct = Len(mxl)
    mxl.dwComponentType = componentType
    rc = mixerGetLineInfo(hmixer, mxl, MIXER_GETLINEINFOF_COMPONENTTYPE)
    If (MMSYSERR_NOERROR = rc) Then
        mxlc.cbStruct = Len(mxlc)
        mxlc.dwLineID = mxl.dwLineID
        mxlc.dwControl = ctrlType
        mxlc.cControls = 1
        mxlc.cbmxctrl = Len(mxc)
        hmem = GlobalAlloc(&H40, Len(mxc))
        mxlc.pamxctrl = GlobalLock(hmem)
        mxc.cbStruct = Len(mxc)
        rc = mixerGetLineControls(hmixer, mxlc, MIXER_GETLINECONTROLSF_ONEBYTYPE)
        If (MMSYSERR_NOERROR = rc) Then
            GetVolumeControl = True
            CopyStructFromPtr mxc, mxlc.pamxctrl, Len(mxc)
        Else
            GetVolumeControl = False
        End If
        GlobalFree (hmem)
        Exit Function
    End If
    GetVolumeControl = False
End Function
'-------------------------------'
'GET VOLUME METER (INPUT/OUTPUT)'
'-------------------------------'
Public Property Get CurrentVolumeMeterOutput() As Long
' Get the current volume output level
    If (fRecording) Then
       For I = 0 To (NUM_BUFFERS - 1)
          If inHdr(I).dwFlags And WHDR_DONE Then
             rc = waveInAddBuffer(hWaveIn, inHdr(I), Len(inHdr(I)))
          End If
       Next
    End If
    mxcd.dwControlID = outputVolCtrl.dwControlID
    mxcd.item = outputVolCtrl.cMultipleItems
    rc = mixerGetControlDetails(hmixer, mxcd, MIXER_GETCONTROLDETAILSF_VALUE)
    CopyStructFromPtr volume, mxcd.paDetails, Len(volume)
    If (volume < 0) Then volume = -volume
    CurrentVolumeMeterOutput = volume
End Property
Public Property Get CurrentVolumeMeterInput() As Long
' Get the current volume input level
    If Not StartInput Then StartInput
    If (fRecording) Then
        For I = 0 To (NUM_BUFFERS - 1)
            If inHdr(I).dwFlags And WHDR_DONE Then
                rc = waveInAddBuffer(hWaveIn, inHdr(I), Len(inHdr(I)))
            End If
        Next
    End If
    mxcd.dwControlID = inputVolCtrl.dwControlID
    mxcd.item = inputVolCtrl.cMultipleItems
    rc = mixerGetControlDetails(hmixer, mxcd, MIXER_GETCONTROLDETAILSF_VALUE)
    CopyStructFromPtr volume, mxcd.paDetails, Len(volume)
    If (volume < 0) Then
        volume = -volume
    End If
    CurrentVolumeMeterInput = volume
End Property
'----------------'
'CLASS PROPERTIES'
'----------------'
Public Property Get MaxVolumeMeterInput() As Long
' Get maximut Volume Input
    MaxVolumeMeterInput = inputVolCtrl.lMaximum
End Property
Public Property Get MinVolumeMeterInput() As Long
' Get minimum Volume Input
    MinVolumeMeterInput = inputVolCtrl.lMinimum
End Property
Public Property Get MaxVolumeMeterOutput() As Long
' Get maximut Volume Output
    MaxVolumeMeterOutput = outputVolCtrl.lMaximum
End Property
Public Property Get MinVolumeMeterOutput() As Long
' Get minimum Volume Output
    MinVolumeMeterOutput = outputVolCtrl.lMinimum
End Property
Public Property Get VolumeMin() As Long
' Get volume Min
    VolumeMin = volCtrl.lMinimum
End Property
Public Property Get WaveMin() As Long
' Get Wave volume min
    WaveMin = wavCtrl.lMinimum
End Property
Public Property Get VolTrebleMin() As Long
' Get Volume Treble min
    VolTrebleMin = volTreb.lMinimum
End Property
Public Property Get WaveTrebleMin() As Long
' Get wave Treble min
    WaveTrebleMin = wavTreb.lMinimum
End Property
Public Property Get VolBassMin() As Long
' Get Volume Bass min
    VolBassMin = volBass.lMinimum
End Property
Public Property Get WaveBassMin() As Long
' Get Wave Bass Min
    WaveBassMin = wavBass.lMinimum
End Property
Public Property Get VolumeMax() As Long
' Get volume bass max
    VolumeMax = volCtrl.lMaximum
End Property
Public Property Get WaveMax() As Long
' Get wave max
    WaveMax = wavCtrl.lMaximum
End Property
Public Property Get VolTrebleMax() As Long
' Get volume treble max
    VolTrebleMax = volTreb.lMaximum
End Property
Public Property Get VolBassMax() As Long
' Get volume bass max
    VolBassMax = volBass.lMaximum
End Property
Public Property Get VolumeLevel() As Long
' Get volume level
    VolumeLevel = GetVolume(volCtrl)
End Property
Public Property Get WaveLevel() As Long
' Get wave level
    WaveLevel = GetVolume(wavCtrl)
End Property
Public Property Get VolumeLevelBass() As Long
' Get volume bass level
    VolumeLevelBass = GetBass(volBass)
End Property
Public Property Get VolumeLevelTreble() As Long
' Get volume treble level
    VolumeLevelTreble = GetTreble(volTreb)
End Property
Public Property Let WaveLevel(ByVal New_Volume As Long)
' Set Wave level
    SetVolume wavCtrl, New_Volume
End Property
Public Property Let VolumeLevel(ByVal New_Volume As Long)
' Set volume level
    SetVolume volCtrl, New_Volume
End Property
Public Property Let WaveLevelBass(ByVal New_Bass As Long)
' Set Wave Bass Level
    SetBass wavBass, New_Bass
End Property
Public Property Let VolumeLevelBass(ByVal New_Bass As Long)
' Set Volume Bass level
    SetBass volBass, New_Bass
End Property
Public Property Let WaveLevelTreble(ByVal New_Treble As Long)
' Set Wave Treble Level
    SetTreble wavTreb, New_Treble
End Property
Public Property Let VolumeLevelTreble(ByVal New_Treble As Long)
' Set Volume Treble level
    SetTreble volTreb, New_Treble
End Property
Public Property Get VolumeMute() As Boolean
' Get Volume Mute Level
    VolumeMute = GetMute(volMute)
End Property
Public Property Get WaveMute() As Boolean
' Get Wave Mute Level
    WaveMute = GetMute(wavMute)
End Property
Public Property Let VolumeMute(New_Mute As Boolean)
' Mute Volume
    SetMute volMute, New_Mute
End Property
Public Property Let WaveMute(New_Mute As Boolean)
' Mute wave sound
    SetMute wavMute, New_Mute
End Property
'---------------'
'TERMINATE CLASS'
'---------------'
Private Sub StopInput()
' Stop Inputing sound
    fRecording = False
    waveInReset hWaveIn
    waveInStop hWaveIn
    For I = 0 To NUM_BUFFERS - 1
        waveInUnprepareHeader hWaveIn, inHdr(I), Len(inHdr(I))
        GlobalFree hmem(I)
    Next
    waveInClose hWaveIn
End Sub
Private Sub Class_Terminate()
' Terminate
    mixerClose hmixer
    If (fRecording = True) Then
        StopInput
    End If
    GlobalFree volHmem
End Sub
