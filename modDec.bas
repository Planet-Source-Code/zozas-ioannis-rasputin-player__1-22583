Attribute VB_Name = "modDec"
'------------'
'DECLARATIONS'
'------------'
Option Explicit
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Long ' To detect sound card
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFY_ICON_DATA) As Boolean
Public Type NOTIFY_ICON_DATA     ' Systray Type
       cbSize As Long
       hwnd As Long
       uId As Long
       uFlags As Long
       uCallBackMessage As Long
       hIcon As Long
       szTip As String * 64
    End Type
' Equalizer Declarations
Public Declare Function waveInAddBuffer Lib "winmm" (ByVal InputDeviceHandle As Long, ByVal WaveHdrPointer As Long, ByVal WaveHdrStructSize As Long) As Long
Public Declare Function waveInPrepareHeader Lib "winmm" (ByVal InputDeviceHandle As Long, ByVal WaveHdrPointer As Long, ByVal WaveHdrStructSize As Long) As Long
Public Declare Function waveInUnprepareHeader Lib "winmm" (ByVal InputDeviceHandle As Long, ByVal WaveHdrPointer As Long, ByVal WaveHdrStructSize As Long) As Long
Public Declare Function waveInGetNumDevs Lib "winmm" () As Long
Public Declare Function waveInGetDevCaps Lib "winmm" Alias "waveInGetDevCapsA" (ByVal uDeviceID As Long, ByVal WaveInCapsPointer As Long, ByVal WaveInCapsStructSize As Long) As Long
Public Declare Function waveInOpen Lib "winmm" (WaveDeviceInputHandle As Long, ByVal WhichDevice As Long, ByVal WaveFormatExPointer As Long, ByVal CallBack As Long, ByVal CallBackInstance As Long, ByVal Flags As Long) As Long
Public Declare Function waveInClose Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long
Public Declare Function waveInStart Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long
Public Declare Function waveInReset Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long
Public Declare Function waveInStop Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long
Public Type WaveFormatEx
    FormatTag As Integer
    Channels As Integer
    SamplesPerSec As Long
    AvgBytesPerSec As Long
    BlockAlign As Integer
    BitsPerSample As Integer
    ExtraDataSize As Integer
End Type
Public Type WAVEHDR
    lpData As Long
    dwBufferLength As Long
    dwBytesRecorded As Long
    dwUser As Long
    dwFlags As Long
    dwLoops As Long
    lpNext As Long
    Reserved As Long
End Type
Public Type WAVEINCAPS
    ManufacturerID As Integer
    ProductID As Integer
    DriverVersion As Long
    ProductName(1 To 32) As Byte
    Formats As Long
    Channels As Integer
    Reserved As Integer
End Type

