Attribute VB_Name = "modDef"
'----------------------------'
'CONSTANTS FOR DEFAULT VALUES'
'----------------------------'
Option Explicit
Public Const DEFAULT_NEXT_SONG = 1          ' Default
Public Const DEFAULT_FF = 10                ' Default
Public Const DEFAULT_RW = 10                ' Default
Public Const DEFAULT_FORM_ON_TOP = False    ' Default
Public Const DEFAULT_AUTOLOAD_LIST = True   ' Default
Public Const DEFAULT_TITLE_SCROLL = True    ' Default
Public Const DEFAULT_FORM_MODE = False      ' Default
Public Const DEFAULT_FORM_TOOLS = False     ' Default
Public Const DEFAULT_MINIMIZE_MODE = False  ' Default
Public Const DEFAULT_AUTOPLAY = False       ' Default
Public Const DEFAULT_TIMER_SPEED = 350      ' Default
Public Const HWND_TOPMOST = -1              ' On top handling
Public Const HWND_NOTOPMOST = -2            ' On top handling
Public Const NIM_ADD = &H0                  ' Systray constant
Public Const NIM_MODIFY = &H1               ' Systray constant
Public Const NIM_DELETE = &H2               ' Systray constant
Public Const NIF_MESSAGE = &H1              ' Systray constant
Public Const NIF_ICON = &H2                 ' Systray constant
Public Const NIF_TIP = &H4                  ' Systray constant
Public Const WM_MOUSEMOVE = &H200           ' Systray constant
Public Const WM_LBUTTONDOWN = &H201         ' Systray constant
Public Const WM_LBUTTONUP = &H202           ' Systray constant
Public Const WM_LBUTTONDBLCLK = &H203       ' Systray constant
Public Const WM_RBUTTONDOWN = &H204         ' Systray constant
Public Const WM_RBUTTONUP = &H205           ' Systray constant
Public Const WM_RBUTTONDBLCLK = &H206       ' Systray constant
' Equalizer definitions
Public Const WAVE_INVALIDFORMAT = &H0&     ' Invalid format
Public Const WAVE_FORMAT_1M08 = &H1&       ' 11.025 kHz, Mono,   8-bit
Public Const WAVE_FORMAT_1S08 = &H2&       ' 11.025 kHz, Stereo, 8-bit
Public Const WAVE_FORMAT_1M16 = &H4&       ' 11.025 kHz, Mono,   16-bit
Public Const WAVE_FORMAT_1S16 = &H8&       ' 11.025 kHz, Stereo, 16-bit
Public Const WAVE_FORMAT_2M08 = &H10&      ' 22.05  kHz, Mono,   8-bit
Public Const WAVE_FORMAT_2S08 = &H20&      ' 22.05  kHz, Stereo, 8-bit
Public Const WAVE_FORMAT_2M16 = &H40&      ' 22.05  kHz, Mono,   16-bit
Public Const WAVE_FORMAT_2S16 = &H80&      ' 22.05  kHz, Stereo, 16-bit
Public Const WAVE_FORMAT_4M08 = &H100&     ' 44.1   kHz, Mono,   8-bit
Public Const WAVE_FORMAT_4S08 = &H200&     ' 44.1   kHz, Stereo, 8-bit
Public Const WAVE_FORMAT_4M16 = &H400&     ' 44.1   kHz, Mono,   16-bit
Public Const WAVE_FORMAT_4S16 = &H800&     ' 44.1   kHz, Stereo, 16-bit
Public Const WAVE_FORMAT_PCM = 1
Public Const WHDR_DONE = &H1&              ' Done bit
Public Const WHDR_PREPARED = &H2&          ' Set if this header has been prepared
Public Const WHDR_BEGINLOOP = &H4&         ' Loop start block
Public Const WHDR_ENDLOOP = &H8&           ' Loop end block
Public Const WHDR_INQUEUE = &H10&          ' Reserved for driver
Public Const WIM_OPEN = &H3BE
Public Const WIM_CLOSE = &H3BF
Public Const WIM_DATA = &H3C0
Public Const ANGLENUMERATOR = 6.283185     ' 2 * Pi
Public Const NUMSAMPLES = 1024             ' Number of Samples
Public Const NUMBITS = 10                  ' Number of Bits

