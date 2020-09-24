Attribute VB_Name = "modVar"
'----------------'
'GLOBAL VARIABLES'
'----------------'
Option Explicit
Public SWP_NOACTIVATE                       ' On top handling
Public SWP_SHOWWINDOW                       ' On top handling
Public PC_FILE_SYSTEM                       ' File System Handling
Public MAINFORM_FLAG                        ' On Top Flag
Public i As Integer                         ' Counter
Public NEXT_SONG As Integer                 ' Next song to play
Public PREVIOUS_X As Integer                ' Previous X position of the form
Public PREVIOUS_Y As Integer                ' Previous Y position of the form
Public PREVIOUS_OPTIONS_X As Integer        ' Previous X position of the options form
Public PREVIOUS_OPTIONS_Y As Integer        ' Previous Y position of the options form
Public PREVIOUS_SCAN_X As Integer           ' Previous X position of the Collection form
Public PREVIOUS_SCAN_Y As Integer           ' Previous Y position of the Collection form
Public FF As Integer                        ' Fast Forward
Public RW As Integer                        ' Rewind
Public TIMER_SPEED As Integer               ' Title scroll speed
Public volume As Integer                    ' Volume
Public SONG As String                       ' Song full path
Public MINUTES_LEFT As String               ' Minutes Left
Public SECONDS_LEFT As String               ' Secods Left
Public RESPONCE As String                   ' Responce string for msgboxes
Public WINDOWS_PATH As String               ' Windows Path
Public FORM_TITLE As String                 ' Scrolling title
Public AUTOLIST As String                   ' Auto Save/Load list filename
Public AUTOPLAY As Boolean                  ' Autoplay song
Public AUTOLOAD_LIST As Boolean             ' Autoload playlist previously used
Public PAUSED_PLAYER As Boolean             ' Player is paused
Public FORM_ON_TOP As Boolean               ' Form stays on top
Public MOVE_FORM As Boolean                 ' Form is moving
Public MOVE_OPTIONS As Boolean              ' Options form is moving
Public MOVE_SCAN As Boolean                 ' Collection form is moving
Public TITLE_SCROLL As Boolean              ' Scroll Title
Public MINIMIZE_MODE As Boolean             ' Minimize to System Tray or Taskbar
Public IS_MINIMIZED As Boolean              ' Form in minimized
Public FORM_MODE As Boolean                 ' Compact or list mode
Public FORM_TOOLS As Boolean                ' Tools shown or not
Public vol As New clsVolume                 ' Mixer declaration
Public SYSTRAYMODE As NOTIFY_ICON_DATA      ' Systray Object to handle
'-------------------'
'TEMPORARY VARIABLES'
'-------------------'
Public EquTemp                              ' Temporary variable for equalizer
Public Equ5temp(4)                          ' Temporary variable for 5 bar amplifier
Public TempVar As Integer                   ' Temporary Variable
Public SecTemp As Integer                   ' Temporary Variable
Public SysTrayTitle As String               ' Temporary Variable for System Tray Mode Title
Public EntryMemo As String                  ' Temporary Variable for list Cut/Copy/Paste Operations
'-------------------'
'EQUALIZER VARIABLES'
'-------------------'
Public DevHandle As Long                    ' Handle of the open audio device
Public Visualizing As Boolean               ' Equalizer is working
Public Divisor As Long                      ' Bar divisor (10.402,5210.4)
Public ScopeHeight As Long                  ' Saves time because hitting up a Long is faster
Public ReversedBits(0 To NUMSAMPLES - 1) As Long ' Bit reservation
Public GREQ_COLOR                           ' Color of GrEq
Public GREQ_ENABLED As Boolean              ' If GrEq is enabled
'-------------------'
'FILE INFO VARIABLES'
'-------------------'
Public BITRATE_LOOKUP(7, 15) As Integer     ' Bit rate
Public ACTUAL_BITRATE As Long               ' Real bit rate
