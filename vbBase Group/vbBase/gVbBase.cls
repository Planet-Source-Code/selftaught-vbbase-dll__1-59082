VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "gVbBase"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'==================================================================================================
'gVbBase.cls                            7/4/04
'
'           LINEAGE:
'               Assembled from several sources:
'                   -Paul Caton's WinsubHook2.tlb from www.planet-source-code.com
'                   -Paul Wilde's vbACOM.dll from www.vbaccelerator.com
'                   -General Api with help from www.allapi.net
'
'           PURPOSE:
'               Provides:
'                   -Global collections of subclasses, windows hooks, vtable subclasses,
'                    api timers, window classes, and windows.
'                   -General unrelated utilities to add to the global namespace.
'
'           CLASSES CREATED BY THIS MODULE:
'               cApiWindows
'               cApiWindowClasses
'               cSubclasses
'               cHooks
'               cTimers
'               cVTableSubclasses
'
'==================================================================================================
'
'1.  Global Collections Interface  - Methods to return the collections maintained by this component.
'2.  Api Utility Interface         - Methods which delegate to useful api calls
'3.  Global Object Interface       - Function that returns a reference to Me so clients may cache a reference
'                                    to this object that is not subject to the 'As New' overhead.


'<General enums from the original WinSubHook2>
Public Enum eMsgWhen
    MSG_AFTER = 1
    MSG_BEFORE = 2
    MSG_BEFORE_AND_AFTER = 3
End Enum

Public Enum eMsg
    ALL_MESSAGES = -1&
    WM_NULL = &H0&
    WM_CREATE = &H1&
    WM_DESTROY = &H2&
    WM_MOVE = &H3&
    WM_SIZE = &H5&
    WM_ACTIVATE = &H6&
    WM_SETFOCUS = &H7&
    WM_KILLFOCUS = &H8&
    WM_ENABLE = &HA&
    WM_SETREDRAW = &HB&
    WM_SETTEXT = &HC&
    WM_GETTEXT = &HD&
    WM_GETTEXTLENGTH = &HE&
    WM_PAINT = &HF&
    WM_CLOSE = &H10&
    WM_QUERYENDSESSION = &H11&
    WM_QUIT = &H12&
    WM_QUERYOPEN = &H13&
    WM_ERASEBKGND = &H14&
    WM_SYSCOLORCHANGE = &H15&
    WM_ENDSESSION = &H16&
    WM_SHOWWINDOW = &H18&
    WM_WININICHANGE = &H1A&
    WM_SETTINGCHANGE = &H1A&
    WM_DEVMODECHANGE = &H1B&
    WM_ACTIVATEAPP = &H1C&
    WM_FONTCHANGE = &H1D&
    WM_TIMECHANGE = &H1E&
    WM_CANCELMODE = &H1F&
    WM_SETCURSOR = &H20&
    WM_MOUSEACTIVATE = &H21&
    WM_CHILDACTIVATE = &H22&
    WM_QUEUESYNC = &H23&
    WM_GETMINMAXINFO = &H24&
    WM_PAINTICON = &H26&
    WM_ICONERASEBKGND = &H27&
    WM_NEXTDLGCTL = &H28&
    WM_SPOOLERSTATUS = &H2A&
    WM_DRAWITEM = &H2B&
    WM_MEASUREITEM = &H2C&
    WM_DELETEITEM = &H2D&
    WM_VKEYTOITEM = &H2E&
    WM_CHARTOITEM = &H2F&
    WM_SETFONT = &H30&
    WM_GETFONT = &H31&
    WM_SETHOTKEY = &H32&
    WM_GETHOTKEY = &H33&
    WM_QUERYDRAGICON = &H37&
    WM_COMPAREITEM = &H39&
    WM_GETOBJECT = &H3D&
    WM_COMPACTING = &H41&
    WM_WINDOWPOSCHANGING = &H46&
    WM_WINDOWPOSCHANGED = &H47&
    WM_POWER = &H48&
    WM_COPYDATA = &H4A&
    WM_CANCELJOURNAL = &H4B&
    WM_NOTIFY = &H4E&
    WM_INPUTLANGCHANGEREQUEST = &H50&
    WM_INPUTLANGCHANGE = &H51&
    WM_TCARD = &H52&
    WM_HELP = &H53&
    WM_USERCHANGED = &H54&
    WM_NOTIFYFORMAT = &H55&
    WM_CONTEXTMENU = &H7B&
    WM_STYLECHANGING = &H7C&
    WM_STYLECHANGED = &H7D&
    WM_DISPLAYCHANGE = &H7E&
    WM_GETICON = &H7F&
    WM_SETICON = &H80&
    WM_NCCREATE = &H81&
    WM_NCDESTROY = &H82&
    WM_NCCALCSIZE = &H83&
    WM_NCHITTEST = &H84&
    WM_NCPAINT = &H85&
    WM_NCACTIVATE = &H86&
    WM_GETDLGCODE = &H87&
    WM_SYNCPAINT = &H88&
    WM_NCMOUSEMOVE = &HA0&
    WM_NCLBUTTONDOWN = &HA1&
    WM_NCLBUTTONUP = &HA2&
    WM_NCLBUTTONDBLCLK = &HA3&
    WM_NCRBUTTONDOWN = &HA4&
    WM_NCRBUTTONUP = &HA5&
    WM_NCRBUTTONDBLCLK = &HA6&
    WM_NCMBUTTONDOWN = &HA7&
    WM_NCMBUTTONUP = &HA8&
    WM_NCMBUTTONDBLCLK = &HA9&
    WM_KEYFIRST = &H100&
    WM_KEYDOWN = &H100&
    WM_KEYUP = &H101&
    WM_CHAR = &H102&
    WM_DEADCHAR = &H103&
    WM_SYSKEYDOWN = &H104&
    WM_SYSKEYUP = &H105&
    WM_SYSCHAR = &H106&
    WM_SYSDEADCHAR = &H107&
    WM_KEYLAST = &H108&
    WM_IME_STARTCOMPOSITION = &H10D&
    WM_IME_ENDCOMPOSITION = &H10E&
    WM_IME_COMPOSITION = &H10F&
    WM_IME_KEYLAST = &H10F&
    WM_INITDIALOG = &H110&
    WM_COMMAND = &H111&
    WM_SYSCOMMAND = &H112&
    WM_TIMER = &H113&
    WM_HSCROLL = &H114&
    WM_VSCROLL = &H115&
    WM_INITMENU = &H116&
    WM_INITMENUPOPUP = &H117&
    WM_MENUSELECT = &H11F&
    WM_MENUCHAR = &H120&
    WM_ENTERIDLE = &H121&
    WM_MENURBUTTONUP = &H122&
    WM_MENUDRAG = &H123&
    WM_MENUGETOBJECT = &H124&
    WM_UNINITMENUPOPUP = &H125&
    WM_MENUCOMMAND = &H126&
    WM_CTLCOLORMSGBOX = &H132&
    WM_CTLCOLOREDIT = &H133&
    WM_CTLCOLORLISTBOX = &H134&
    WM_CTLCOLORBTN = &H135&
    WM_CTLCOLORDLG = &H136&
    WM_CTLCOLORSCROLLBAR = &H137&
    WM_CTLCOLORSTATIC = &H138&
    WM_MOUSEFIRST = &H200&
    WM_MOUSEMOVE = &H200&
    WM_LBUTTONDOWN = &H201&
    WM_LBUTTONUP = &H202&
    WM_LBUTTONDBLCLK = &H203&
    WM_RBUTTONDOWN = &H204&
    WM_RBUTTONUP = &H205&
    WM_RBUTTONDBLCLK = &H206&
    WM_MBUTTONDOWN = &H207&
    WM_MBUTTONUP = &H208&
    WM_MBUTTONDBLCLK = &H209&
    WM_MOUSEWHEEL = &H20A&
    WM_PARENTNOTIFY = &H210&
    WM_ENTERMENULOOP = &H211&
    WM_EXITMENULOOP = &H212&
    WM_NEXTMENU = &H213&
    WM_SIZING = &H214&
    WM_CAPTURECHANGED = &H215&
    WM_MOVING = &H216&
    WM_DEVICECHANGE = &H219&
    WM_MDICREATE = &H220&
    WM_MDIDESTROY = &H221&
    WM_MDIACTIVATE = &H222&
    WM_MDIRESTORE = &H223&
    WM_MDINEXT = &H224&
    WM_MDIMAXIMIZE = &H225&
    WM_MDITILE = &H226&
    WM_MDICASCADE = &H227&
    WM_MDIICONARRANGE = &H228&
    WM_MDIGETACTIVE = &H229&
    WM_MDISETMENU = &H230&
    WM_ENTERSIZEMOVE = &H231&
    WM_EXITSIZEMOVE = &H232&
    WM_DROPFILES = &H233&
    WM_MDIREFRESHMENU = &H234&
    WM_IME_SETCONTEXT = &H281&
    WM_IME_NOTIFY = &H282&
    WM_IME_CONTROL = &H283&
    WM_IME_COMPOSITIONFULL = &H284&
    WM_IME_SELECT = &H285&
    WM_IME_CHAR = &H286&
    WM_IME_REQUEST = &H288&
    WM_IME_KEYDOWN = &H290&
    WM_IME_KEYUP = &H291&
    WM_MOUSEHOVER = &H2A1&
    WM_MOUSELEAVE = &H2A3&
    WM_CUT = &H300&
    WM_COPY = &H301&
    WM_PASTE = &H302&
    WM_CLEAR = &H303&
    WM_UNDO = &H304&
    WM_RENDERFORMAT = &H305&
    WM_RENDERALLFORMATS = &H306&
    WM_DESTROYCLIPBOARD = &H307&
    WM_DRAWCLIPBOARD = &H308&
    WM_PAINTCLIPBOARD = &H309&
    WM_VSCROLLCLIPBOARD = &H30A&
    WM_SIZECLIPBOARD = &H30B&
    WM_ASKCBFORMATNAME = &H30C&
    WM_CHANGECBCHAIN = &H30D&
    WM_HSCROLLCLIPBOARD = &H30E&
    WM_QUERYNEWPALETTE = &H30F&
    WM_PALETTEISCHANGING = &H310&
    WM_PALETTECHANGED = &H311&
    WM_HOTKEY = &H312&
    WM_PRINT = &H317&
    WM_PRINTCLIENT = &H318&
    WM_THEMECHANGED = &H31A&
    WM_HANDHELDFIRST = &H358&
    WM_HANDHELDLAST = &H35F&
    WM_AFXFIRST = &H360&
    WM_AFXLAST = &H37F&
    WM_PENWINFIRST = &H380&
    WM_PENWINLAST = &H38F&
    WM_USER = &H400&
    WM_APP = &H8000&
    WM_UPDATEUISTATE = &H128&
    WM_CHANGEUISTATE = &H127&

End Enum

Public Enum eClassStyle
    CS_VREDRAW = 1
    CS_HREDRAW = 2
    CS_KEYCVTWINDOW = 4
    CS_DBLCLKS = 8
    CS_OWNDC = 32
    CS_CLASSDC = 64
    CS_PARENTDC = 128
    CS_NOKEYCVT = 25
    CS_NOCLOSE = 512
    CS_SAVEBITS = 2048
    CS_BYTEALIGNCLIENT = 4096
    CS_BYTEALIGNWINDOW = 8192
    CS_PUBLICCLASS = 16384
End Enum

Public Enum eWindowClass
    AS_WINDOWCLASS = 0
    PREDEFINED_BUTTON = 1
    PREDEFINED_COMBOBOX = 2
    PREDEFINED_EDIT = 3
    PREDEFINED_LISTBOX = 4
    PREDEFINED_MDICLIENT = 5
    PREDEFINED_RICHEDIT = 6
    PREDEFINED_RICHEDIT_CLASS = 7
    PREDEFINED_SCROLLBAR = 8
    PREDEFINED_STATIC = 9
End Enum

Public Enum eWindowStyle
    WS_OVERLAPPED = &H0
    WS_POPUP = &H80000000
    WS_CHILD = &H40000000
    WS_MINIMIZE = &H20000000
    WS_VISIBLE = &H10000000
    WS_DISABLED = &H8000000
    WS_CLIPSIBLINGS = &H4000000
    WS_CLIPCHILDREN = &H2000000
    WS_MAXIMIZE = &H1000000
    WS_CAPTION = &HC00000
    WS_BORDER = &H800000
    WS_DLGFRAME = &H400000
    WS_VSCROLL = &H200000
    WS_HSCROLL = &H100000
    WS_SYSMENU = &H80000
    WS_THICKFRAME = &H40000
    WS_GROUP = &H20000
    WS_TABSTOP = &H10000
    WS_MINIMIZEBOX = &H20000
    WS_MAXIMIZEBOX = &H10000
    WS_TILED = &H0
    WS_ICONIC = &H20000000
    WS_SIZEBOX = &H40000
    WS_OVERLAPPEDWINDOW = &HCF0000
    WS_POPUPWINDOW = &H80880000
End Enum

Public Enum eWindowStyleEx
    WS_EX_DLGMODALFRAME = &H1
    WS_EX_NOPARENTNOTIFY = &H4
    WS_EX_TOPMOST = &H8
    WS_EX_ACCEPTFILES = &H10
    WS_EX_TRANSPARENT = &H20
    WS_EX_MDICHILD = &H40
    WS_EX_TOOLWINDOW = &H80
    WS_EX_WINDOWEDGE = &H100
    WS_EX_CLIENTEDGE = &H200
    WS_EX_CONTEXTHELP = &H400
    WS_EX_RIGHT = &H1000
    WS_EX_LEFT = &H0
    WS_EX_RTLREADING = &H2000
    WS_EX_LTRREADING = &H0
    WS_EX_LEFTSCROLLBAR = &H4000
    WS_EX_RIGHTSCROLLBAR = &H0
    WS_EX_CONTROLPARENT = &H10000
    WS_EX_STATICEDGE = &H20000
    WS_EX_APPWINDOW = &H40000
    WS_EX_OVERLAPPEDWINDOW = &H300
    WS_EX_PALETTEWINDOW = &H188
    WS_EX_LAYERED = &H80000
    WS_EX_COMPOSITED = &H2000000
    WS_EX_NOACTIVATE = &H8000000
End Enum
 
Public Enum eHookCode
    HCBT_MOVESIZE = 0
    HCBT_MINMAX = 1
    HCBT_QS = 2
    HCBT_CREATEWND = 3
    HCBT_DESTROYWND = 4
    HCBT_ACTIVATE = 5
    HCBT_CLICKSKIPPED = 6
    HCBT_KEYSKIPPED = 7
    HCBT_SYSCOMMAND = 8
    HCBT_SETFOCUS = 9
    PM_NOREMOVE = 0
    PM_REMOVE = 1
    PM_NOYIELD = 2
    HC_ACTION = 0
    HC_GETNEXT = 1
    HC_SKIP = 2
    HC_NOREMOVE = 3
    HC_NOREM = 3
    HC_SYSMODALON = 4
    HC_SYSMODALOFF = 5
    MSGF_DIALOGBOX = 0
    MSGF_MESSAGEBOX = 1
    MSGF_MENU = 2
    MSGF_SCROLLBAR = 5
    MSGF_NEXTWINDOW = 6
    MSGF_MAX = 8
    MSGF_USER = 4096
    MSGF_DDEMGR = 32769
    HSHELL_WINDOWCREATED = 1
    HSHELL_WINDOWDESTROYED = 2
    HSHELL_ACTIVATESHELLWINDOW = 3
    HSHELL_WINDOWACTIVATED = 4
    HSHELL_GETMINRECT = 5
    HSHELL_REDRAW = 6
    HSHELL_TASKMAN = 7
    HSHELL_LANGUAGE = 8
End Enum

Public Enum eHookType
    WH_MSGFILTER = -1
    WH_JOURNALRECORD = 0
    WH_JOURNALPLAYBACK = 1
    WH_KEYBOARD = 2
    WH_GETMESSAGE = 3
    WH_CALLWNDPROC = 4
    WH_CBT = 5
    WH_SYSMSGFILTER = 6
    WH_MOUSE = 7
    WH_DEBUG = 9
    WH_SHELL = 10
    WH_FOREGROUNDIDLE = 11
    WH_CALLWNDPROCRET = 12
    WH_KEYBOARD_LL = 13
    WH_MOUSE_LL = 14
End Enum
'</General enums from the original WinSubHook2>


'<General types from the original WinSubHook2>
Public Type tRect
    Left            As Long
    Top             As Long
    Right           As Long
    Bottom          As Long
End Type

Public Type tRectSize
    Left            As Long
    Top             As Long
    Width           As Long
    Height          As Long
End Type

Public Type tVBBasePoint
    x               As Long
    y               As Long
End Type

Public Type tSize
    cx              As Long
    cy              As Long
End Type

Public Type tMsg
    hWnd            As Long
    Message         As Long
    wParam          As Long
    lParam          As Long
    Time            As Long
    Pt              As tVBBasePoint
End Type


Public Type tCWPStruct
    lParam          As Long
    wParam          As Long
    Message         As Long
    hWnd            As Long
End Type

Public Type tCWPRetStruct
    lResult         As Long
    lParam          As Long
    wParam          As Long
    Message         As Long
    hWnd            As Long
End Type

Public Type tCBTActivateStruct
    fMouse          As Long
    hWndActive      As Long
End Type

Public Type tCreateStruct
    lpCreateParams  As Long
    hInstance       As Long
    hMenu           As Long
    hWndParent      As Long
    cy              As Long
    cx              As Long
    y               As Long
    x               As Long
    Style           As eWindowStyle
    lpszName        As Long
    lpszClass       As Long
    dwExStyle       As eWindowStyleEx
End Type

Public Type tCBTCreateWndStruct
    lpcs            As Long
    hWndInsertAfter As Long
End Type

Public Type tMouseHookStruct
    Pt              As tVBBasePoint
    hWnd            As Long
    wHitTestCode    As Long
    dwExtraInfo     As Long
End Type

Public Type tDebugHookInfo
    idThread        As Long
    idThreadInstaller As Long
    lParam          As Long
    wParam          As Long
    Code            As Integer
End Type

Public Type tEventMsg
    Message         As Long
    ParamL          As Long
    ParamH          As Long
    Time            As Long
    hWnd            As Long
End Type
    
Public Type tKBDLLHookStruct
    VKCode          As Long
    ScanCode        As Long
    Flags           As Long
    Time            As Long
    dwExtraInfo     As Long
End Type
    
Public Type tMSLLHookStruct
    Pt              As tVBBasePoint
    MouseData       As Long
    Flags           As Long
    Time            As Long
    dwExtraInfo     As Long
End Type
'</General types from the original WinSubHook2>

'<Associated with mWindow's General Api Interface>
Public Enum eSWCmd
    SW_HIDE = 0
    SW_MAXIMIZE = 3
    SW_MINIMIZE = 6
    SW_RESTORE = 9
    SW_SHOW = 5
    SW_SHOWDEFAULT = 10
    SW_SHOWMAXIMIZED = 3
    SW_SHOWMINIMIZED = 2
    SW_SHOWMINNOACTIVE = 7
    SW_SHOWNA = 8
    SW_SHOWNOACTIVATE = 4
    SW_SHOWNORMAL = 1
End Enum

Public Enum eSWPhWnd
    HWND_BOTTOM = 1
    HWND_NOTOPMOST = -2
    HWND_TOP = 0
    HWND_TOPMOST = -1
End Enum

Public Enum eSWPFlags
    SWP_ASYNCWINDOWPOS = &H4000
    SWP_DEFERERASE = &H2000
    SWP_FRAMECHANGED = &H20
    SWP_HIDEWINDOW = &H80
    SWP_NOACTIVATE = &H10
    SWP_NOCOPYBITS = &H100
    SWP_NOMOVE = &H2
    SWP_NOOWNERZORDER = &H200
    SWP_NOREDRAW = &H8
    SWP_NOSENDCHANGING = &H400
    SWP_NOSIZE = &H1
    SWP_NOZORDER = &H4
    SWP_SHOWWINDOW = &H40
End Enum

Public Enum eRedrawFlags
    RDW_ALLCHILDREN = &H80
    RDW_ERASE = &H4
    RDW_ERASENOW = &H200
    RDW_FRAME = &H400
    RDW_INTERNALPAINT = &H2
    RDW_INVALIDATE = &H1
    RDW_NOCHILDREN = &H40
    RDW_NOERASE = &H20
    RDW_NOFRAME = &H800
    RDW_NOINTERNALPAINT = &H10
    RDW_UPDATENOW = &H100
    RDW_VALIDATE = &H8
End Enum

'</Associated with mWindow's General Api Interface>

'<Associated with cOleHooks>

'standard Dispatch ID constants
Private Const DISPID_UNKNOWN As Long = (-1&)

Private Const DISPID_AMBIENT_BACKCOLOR As Long = (-701&)
Private Const DISPID_AMBIENT_DISPLAYNAME As Long = (-702&)
Private Const DISPID_AMBIENT_FONT As Long = (-703&)
Private Const DISPID_AMBIENT_FORECOLOR As Long = (-704&)
Private Const DISPID_AMBIENT_LOCALEID As Long = (-705&)
Private Const DISPID_AMBIENT_MESSAGEREFLECT As Long = (-706&)
Private Const DISPID_AMBIENT_SCALEUNITS As Long = (-707&)
Private Const DISPID_AMBIENT_TEXTALIGN As Long = (-708&)
Private Const DISPID_AMBIENT_USERMODE As Long = (-709&)
Private Const DISPID_AMBIENT_UIDEAD As Long = (-710&)
Private Const DISPID_AMBIENT_SHOWGRABHANDLES As Long = (-711&)
Private Const DISPID_AMBIENT_SHOWHATCHING As Long = (-712&)
Private Const DISPID_AMBIENT_DISPLAYASDEFAULT As Long = (-713&)
Private Const DISPID_AMBIENT_SUPPORTSMNEMONICS As Long = (-714&)
Private Const DISPID_AMBIENT_AUTOCLIP As Long = (-715&)
Private Const DISPID_AMBIENT_APPEARANCE As Long = (-716&)

Private Const DISPID_AMBIENT_CODEPAGE As Long = (-725&)
Private Const DISPID_AMBIENT_PALETTE As Long = (-726&)
Private Const DISPID_AMBIENT_CHARSET As Long = (-727&)
Private Const DISPID_AMBIENT_TRANSFERPRIORITY As Long = (-728&)

Private Const DISPID_AMBIENT_RIGHTTOLEFT As Long = (-732&)
Private Const DISPID_AMBIENT_TOPTOBOTTOM As Long = (-733&)

Private Const DISPID_Name As Long = (-800&)
Private Const DISPID_Delete As Long = (-801&)
Private Const DISPID_Object As Long = (-802&)
Private Const DISPID_Parent As Long = (-803&)


Public Enum eAmbientDispID
    vbAmbientDispIDUnknown = DISPID_UNKNOWN
    
    vbAmbientDispIDBackColor = DISPID_AMBIENT_BACKCOLOR
    vbAmbientDispIDDisplayName = DISPID_AMBIENT_DISPLAYNAME
    vbAmbientDispIDFont = DISPID_AMBIENT_FONT
    vbAmbientDispIDForeColor = DISPID_AMBIENT_FORECOLOR
    vbAmbientDispIDLocaleID = DISPID_AMBIENT_LOCALEID
    vbAmbientDispIDMessageReflect = DISPID_AMBIENT_MESSAGEREFLECT
    vbAmbientDispIDScaleUnits = DISPID_AMBIENT_SCALEUNITS
    vbAmbientDispIDTextAlign = DISPID_AMBIENT_TEXTALIGN
    vbAmbientDispIDUserMode = DISPID_AMBIENT_USERMODE
    vbAmbientDispIDUIDead = DISPID_AMBIENT_UIDEAD
    vbAmbientDispIDShowGrabHandles = DISPID_AMBIENT_SHOWGRABHANDLES
    vbAmbientDispIDShowHatching = DISPID_AMBIENT_SHOWHATCHING
    vbAmbientDispIDDisplayAsDefault = DISPID_AMBIENT_DISPLAYASDEFAULT
    vbAmbientDispIDSupportsMnemonics = DISPID_AMBIENT_SUPPORTSMNEMONICS
    vbAmbientDispIDAutoClip = DISPID_AMBIENT_AUTOCLIP
    vbAmbientDispIDAppearance = DISPID_AMBIENT_APPEARANCE

    vbAmbientDispIDCodePage = DISPID_AMBIENT_CODEPAGE
    vbAmbientDispIDPalette = DISPID_AMBIENT_PALETTE
    vbAmbientDispIDCharSet = DISPID_AMBIENT_CHARSET
    vbAmbientDispIDTransferPriority = DISPID_AMBIENT_TRANSFERPRIORITY
    
    vbAmbientDispIDRightToLeft = DISPID_AMBIENT_RIGHTTOLEFT
    vbAmbientDispIDTopToBottom = DISPID_AMBIENT_TOPTOBOTTOM

End Enum

Public Enum eControlInfoFlags
    ctlInfoEatsReturn = CTRLINFO_EATS_RETURN
    ctlInfoEatsEscape = CTRLINFO_EATS_ESCAPE
End Enum
'</Associated with cOleHooks>

'<Associated with the Api Utility Interface>
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetCursorPos Lib "user32" (lpPoint As tVBBasePoint) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Public Enum eSysColor
    COLOR_ACTIVEBORDER = 10
    COLOR_ACTIVECAPTION = 2
    COLOR_ADJ_MAX = 100
    COLOR_ADJ_MIN = -100 'shorts
    COLOR_APPWORKSPACE = 12
    COLOR_BACKGROUND = 1
    COLOR_BTNFACE = 15
    COLOR_BTNHIGHLIGHT = 20
    COLOR_BTNSHADOW = 16
    COLOR_BTNTEXT = 18
    COLOR_CAPTIONTEXT = 9
    COLOR_GRAYTEXT = 17
    COLOR_HIGHLIGHT = 13
    COLOR_HIGHLIGHTTEXT = 14
    COLOR_INACTIVEBORDER = 11
    COLOR_INACTIVECAPTION = 3
    COLOR_INACTIVECAPTIONTEXT = 19
    COLOR_MENU = 4
    COLOR_MENUTEXT = 7
    COLOR_SCROLLBAR = 0
    COLOR_WINDOW = 5
    COLOR_WINDOWFRAME = 6
    COLOR_WINDOWTEXT = 8
End Enum

Public Enum eSysMetric
    SM_CMETRICS = 44
    SM_CMOUSEBUTTONS = 43
    SM_CXBORDER = 5
    SM_CXCURSOR = 13
    SM_CXDLGFRAME = 7
    SM_CXDOUBLECLK = 36
    SM_CXFRAME = 32
    SM_CXFULLSCREEN = 16
    SM_CXHSCROLL = 21
    SM_CXHTHUMB = 10
    SM_CXICON = 11
    SM_CXICONSPACING = 38
    SM_CXMIN = 28
    SM_CXMINTRACK = 34
    SM_CXSCREEN = 0
    SM_CXSIZE = 30
    SM_CXVSCROLL = 2
    SM_CYBORDER = 6
    SM_CYCAPTION = 4
    SM_CYCURSOR = 14
    SM_CYDLGFRAME = 8
    SM_CYDOUBLECLK = 37
    SM_CYFRAME = 33
    SM_CYFULLSCREEN = 17
    SM_CYHSCROLL = 3
    SM_CYICON = 12
    SM_CYICONSPACING = 39
    SM_CYKANJIWINDOW = 18
    SM_CYMENU = 15
    SM_CYMIN = 29
    SM_CYMINTRACK = 35
    SM_CYSCREEN = 1
    SM_CYSIZE = 31
    SM_CYVSCROLL = 20
    SM_CYVTHUMB = 9
    SM_CXMENUCHECK = 71
End Enum

Public Enum eVirtualKeys
   
    VK_LBUTTON = &H1
    VK_RBUTTON = &H2
    VK_CANCEL = &H3
    VK_MBUTTON = &H4           '/* NOT contiguous with L & RBUTTON */

    VK_BACK = &H8
    VK_TAB = &H9

    VK_CLEAR = &HC
    VK_RETURN = &HD

    VK_SHIFT = &H10
    VK_CONTROL = &H11
    VK_MENU = &H12
    VK_PAUSE = &H13
    VK_CAPITAL = &H14

    VK_KANA = &H15
    'VK_HANGEUL = &H15         '/* old name - should be here for compatibility */
    VK_HANGUL = &H15
    VK_JUNJA = &H17
    VK_FINAL = &H18
    VK_HANJA = &H19
    VK_KANJI = &H19

    VK_ESCAPE = &H1B

    VK_CONVERT = &H1C
    VK_NONCONVERT = &H1D
    VK_ACCEPT = &H1E
    VK_MODECHANGE = &H1F

    VK_SPACE = &H20
    VK_PRIOR = &H21
    VK_NEXT = &H22
    VK_END = &H23
    VK_HOME = &H24
    VK_LEFT = &H25
    VK_UP = &H26
    VK_RIGHT = &H27
    VK_DOWN = &H28
    VK_SELECT = &H29
    VK_PRINT = &H2A
    VK_EXECUTE = &H2B
    VK_SNAPSHOT = &H2C
    VK_INSERT = &H2D
    VK_DELETE = &H2E
    VK_HELP = &H2F

    VK_A = vbKeyA
    VK_B = vbKeyB
    VK_C = vbKeyC
    VK_D = vbKeyD
    VK_E = vbKeyE
    VK_F = vbKeyF
    VK_G = vbKeyG
    VK_H = vbKeyH
    VK_I = vbKeyI
    VK_J = vbKeyJ
    VK_K = vbKeyK
    VK_L = vbKeyL
    VK_M = vbKeyM
    VK_N = vbKeyN
    VK_O = vbKeyO
    VK_P = vbKeyP
    VK_Q = vbKeyQ
    VK_R = vbKeyR
    VK_S = vbKeyS
    VK_T = vbKeyT
    VK_U = vbKeyU
    VK_V = vbKeyV
    VK_W = vbKeyW
    VK_X = vbKeyX
    VK_Y = vbKeyY
    VK_Z = vbKeyZ
    
    VK_0 = vbKey0
    VK_1 = vbKey1
    VK_2 = vbKey2
    VK_3 = vbKey3
    VK_4 = vbKey4
    VK_5 = vbKey5
    VK_6 = vbKey6
    VK_7 = vbKey7
    VK_8 = vbKey8
    VK_9 = vbKey9

    'VK_LWIN = &H5B
    'VK_RWIN = &H5C
    'VK_APPS = &H5D

    VK_NUMPAD0 = &H60
    VK_NUMPAD1 = &H61
    VK_NUMPAD2 = &H62
    VK_NUMPAD3 = &H63
    VK_NUMPAD4 = &H64
    VK_NUMPAD5 = &H65
    VK_NUMPAD6 = &H66
    VK_NUMPAD7 = &H67
    VK_NUMPAD8 = &H68
    VK_NUMPAD9 = &H69
    VK_MULTIPLY = &H6A
    VK_ADD = &H6B
    VK_SEPARATOR = &H6C
    VK_SUBTRACT = &H6D
    VK_DECIMAL = &H6E
    VK_DIVIDE = &H6F
    VK_F1 = &H70
    VK_F2 = &H71
    VK_F3 = &H72
    VK_F4 = &H73
    VK_F5 = &H74
    VK_F6 = &H75
    VK_F7 = &H76
    VK_F8 = &H77
    VK_F9 = &H78
    VK_F10 = &H79
    VK_F11 = &H7A
    VK_F12 = &H7B
    VK_F13 = &H7C
    VK_F14 = &H7D
    VK_F15 = &H7E
    VK_F16 = &H7F
    VK_F17 = &H80
    VK_F18 = &H81
    VK_F19 = &H82
    VK_F20 = &H83
    VK_F21 = &H84
    VK_F22 = &H85
    VK_F23 = &H86
    VK_F24 = &H87

    VK_NUMLOCK = &H90
    VK_SCROLL = &H91
   
'/*
' * VK_L* & VK_R* - left and right Alt, Ctrl and Shift virtual keys.
' * Used only as parameters to GetAsyncKeyState() and GetKeyState().
' * No other API or message will distinguish left and right keys in this way.
' */
   
    VK_LSHIFT = &HA0
    VK_RSHIFT = &HA1
    VK_LCONTROL = &HA2
    VK_RCONTROL = &HA3
    VK_LMENU = &HA4
    VK_RMENU = &HA5
   
    
    
    VK_LWIN = &H5B 'Left Windows key (Microsoft® Natural® keyboard)
    VK_RWIN = &H5C 'Right Windows key (Natural keyboard)
    VK_APPS = &H5D 'Applications key (Natural keyboard)
    VK_SLEEP = &H5F 'Computer Sleep key
   
    VK_BROWSER_BACK = &HA6 'Windows 2000/XP: Browser Back key
    VK_BROWSER_FORWARD = &HA7 'Windows 2000/XP: Browser Forward key
    VK_BROWSER_REFRESH = &HA8 'Windows 2000/XP: Browser Refresh key
    VK_BROWSER_STOP = &HA9 'Windows 2000/XP: Browser Stop key
    VK_BROWSER_SEARCH = &HAA 'Windows 2000/XP: Browser Search key
    VK_BROWSER_FAVORITES = &HAB 'Windows 2000/XP: Browser Favorites key
    VK_BROWSER_HOME = &HAC 'Windows 2000/XP: Browser Start and Home key
    VK_VOLUME_MUTE = &HAD 'Windows 2000/XP: Volume Mute key
    VK_VOLUME_DOWN = &HAE  'Windows 2000/XP: Volume Down key
    VK_VOLUME_UP = &HAF  'Windows 2000/XP: Volume Up key
    VK_MEDIA_NEXT_TRACK = &HB0  'Windows 2000/XP: Next Track key
    VK_MEDIA_PREV_TRACK = &HB1  'Windows 2000/XP: Previous Track key
    VK_MEDIA_STOP = &HB2  'Windows 2000/XP: Stop Media key
    VK_MEDIA_PLAY_PAUSE = &HB3  'Windows 2000/XP: Play/Pause Media key
    VK_LAUNCH_MAIL = &HB4  'Windows 2000/XP: Start Mail key
    VK_LAUNCH_MEDIA_SELECT = &HB5  'Windows 2000/XP: Select Media key
    VK_LAUNCH_APP1 = &HB6  'Windows 2000/XP: Start Application 1 key
    VK_LAUNCH_APP2 = &HB7  'Windows 2000/XP: Start Application 2 key
    VK_OEM_1 = &HBA 'Used for miscellaneous characters; it can vary by keyboard. Windows 2000/XP: For the US standard keyboard, the ';:' key
    
    VK_OEM_PLUS = &HBB 'Windows 2000/XP: For any country/region, the '+' key
    VK_OEM_COMMA = &HBC 'Windows 2000/XP: For any country/region, the ',' key
    VK_OEM_MINUS = &HBD 'Windows 2000/XP: For any country/region, the '-' key
    VK_OEM_PERIOD = &HBE 'Windows 2000/XP: For any country/region, the '.' key
    VK_OEM_2 = &HBF 'Used for miscellaneous characters; it can vary by keyboard. Windows 2000/XP: For the US standard keyboard, the '/?' key
    VK_OEM_3 = &HC0 'Used for miscellaneous characters; it can vary by keyboard. Windows 2000/XP: For the US standard keyboard, the '`~' key
    
    '—  C1–D7 Reserved
    '—  D8–DA Unassigned
    VK_OEM_4 = &HDB 'Used for miscellaneous characters; it can vary by keyboard. Windows 2000/XP: For the US standard keyboard, the '[{' key
    VK_OEM_5 = &HDC 'Used for miscellaneous characters; it can vary by keyboard. Windows 2000/XP: For the US standard keyboard, the '\|' key
    VK_OEM_6 = &HDD 'Used for miscellaneous characters; it can vary by keyboard Windows 2000/XP: For the US standard keyboard, the ']}' key
    VK_OEM_7 = &HDE ' Used for miscellaneous characters; it can vary by keyboard. Windows 2000/XP: For the US standard keyboard, the 'single-quote/double-quote' key
    VK_OEM_8 = &HDF 'Used for miscellaneous characters; it can vary by keyboard. —  E0 Reserved
    '- E1 OEM specific
    VK_OEM_102 = &HE2 'Windows 2000/XP: Either the angle bracket key or the backslash key on the RT 102-key keyboard
    ' E3–E4 OEM specific
    VK_PROCESSKEY = &HE5 'Windows 95/98/Me, Windows NT 4.0, Windows 2000/XP: IME PROCESS key
    ' E6 OEM specific
    VK_PACKET = &HE7 'Windows 2000/XP: Used to pass Unicode characters as if they were keystrokes. The VK_PACKET key is the low word of a 32-bit Virtual Key value used for non-keyboard input methods. For more information, see Remark in KEYBDINPUT, SendInput, WM_KEYDOWN, and WM_KEYUP
    '—  E8 Unassigned
    ' E9–F5 OEM specific
    VK_ATTN = &HF6 'Attn key
    VK_CRSEL = &HF7 'CrSel key
    VK_EXSEL = &HF8 'ExSel key
    VK_EREOF = &HF9 'Erase EOF key
    VK_PLAY = &HFA 'Play key
    VK_ZOOM = &HFB 'Zoom key
    VK_NONAME = &HFC 'Reserved for future use
    VK_PA1 = &HFD 'PA1 key
    VK_OEM_CLEAR = &HFE 'Clear key
End Enum
'</Associated with the Api Utility Interface>

Public Enum eVbBaseErrors
    vbbInvalidProcedureCall = 5
    vbbTypeMismatch = 13
    vbbItemDetached = 35605
    vbbCollChangedDuringEnum = 35606
    vbbKeyNotFound = 35601
    vbbKeyAlreadyExists = 35602
    vbbApiError = 47822
End Enum


'<Private Interface>
'</Private Interface>

'<Friendly Interface>
'</Friendly Interface>

'<Public Interface>
'<Global Collections Interface>
Public Function Subclasses( _
            ByVal oWho As iSubclass) _
                As cSubclasses
    If Not oWho Is Nothing Then
        Set Subclasses = New cSubclasses
        Subclasses.fInit ObjPtr(oWho)
    Else
        gErr vbbInvalidProcedureCall, "gVbBase.Subclasses"
    End If
End Function

Public Function ApiWindows( _
            ByVal oWho As Object) _
                As cApiWindows
    If Not oWho Is Nothing Then
        Dim ltPointer As tPointer
        ltPointer.iIndex = -1&
        ltPointer.iId = ObjPtr(oWho)
        Set ApiWindows = New cApiWindows
        ApiWindows.fInit ltPointer
    Else
        gErr vbbInvalidProcedureCall, "gVbBase.ApiWindows"
    End If
End Function

Public Function ApiWindowClasses() _
                As cApiWindowClasses
    Set ApiWindowClasses = New cApiWindowClasses
End Function

Public Function Timers( _
            ByVal oWho As iTimer) _
                As cTimers
    If Not oWho Is Nothing Then
        Set Timers = New cTimers
        Timers.fInit ObjPtr(oWho)
    Else
        gErr vbbInvalidProcedureCall, "gVbBase.Timers"
    End If
End Function

Public Function Hooks( _
            ByVal oWho As iHook) _
                As cHooks
    If Not oWho Is Nothing Then
        Set Hooks = New cHooks
        Hooks.fInit ObjPtr(oWho)
    Else
        gErr vbbInvalidProcedureCall, "gVbBase.Hooks"
    End If
End Function

Public Function VTableSubclasses() _
                As cVTableSubclasses
    On Error Resume Next

    Set VTableSubclasses = New cVTableSubclasses

End Function

'<Global Collections Interface>


'<Api Utility Interface>
Public Function xKeyIsDown( _
            ByVal iVirtKey As eVirtualKeys, _
   Optional ByVal bAsync As Boolean = True) _
                As Boolean
    If bAsync _
        Then xKeyIsDown = CBool(GetAsyncKeyState(iVirtKey) And &H8000) _
        Else xKeyIsDown = CBool(GetKeyState(iVirtKey) And &H8000)
        
End Function

Public Function xKeyIsToggled( _
            ByVal iVirtKey As eVirtualKeys) _
                As Boolean
    xKeyIsToggled = CBool(GetKeyState(iVirtKey) And 1)
End Function

Public Function xActivateIPAO( _
            ByVal oObject As Object) _
                As Boolean
      On Error Resume Next
      
    Dim pOleObject                  As vbBaseTlb.IOleObject
    Dim pOleInPlaceSite             As vbBaseTlb.IOleInPlaceSite
    Dim pOleInPlaceFrame            As vbBaseTlb.IOleInPlaceFrame
    Dim pOleInPlaceUIWindow         As vbBaseTlb.IOleInPlaceUIWindow
    Dim pOleInPlaceActiveObject     As vbBaseTlb.IOleInPlaceActiveObject
    Dim PosRect                     As vbBaseTlb.RECT
    Dim ClipRect                    As vbBaseTlb.RECT
    Dim FrameInfo                   As vbBaseTlb.OLEINPLACEFRAMEINFO
    
    Set pOleObject = oObject
    Set pOleInPlaceActiveObject = oObject
    
    pOleObject.GetClientSite pOleInPlaceSite
    pOleInPlaceSite.GetWindowContext pOleInPlaceFrame, pOleInPlaceUIWindow, PosRect, ClipRect, FrameInfo
    pOleInPlaceFrame.SetActiveObject pOleInPlaceActiveObject, vbNullString
    If Not pOleInPlaceUIWindow Is Nothing _
        Then pOleInPlaceUIWindow.SetActiveObject pOleInPlaceActiveObject, vbNullString

End Function

Public Function xApiWindow( _
            ByVal hWnd As Long) _
                As cApiWindow
    If IsWindow(hWnd) Then
        Set xApiWindow = New cApiWindow
        xApiWindow.fInit hWnd, WindowClassName(hWnd)
    End If
End Function

Public Function xOnControlInfoChanged( _
            ByVal oControl As Object _
                ) As Boolean
    Dim loOleObject As IOleObject
    Dim loClientSite As IOleClientSite
    Dim loUnknown As stdole.IUnknown
    Dim loControlSite As IOleControlSite
    
    On Error Resume Next
    
    Set loOleObject = oControl
    loOleObject.GetClientSite loClientSite
    Set loUnknown = loClientSite
    Set loControlSite = loUnknown
    loControlSite.OnControlInfoChanged
    
    xOnControlInfoChanged = CBool(Err.Number)
    On Error GoTo 0
End Function

Public Function xCursorPos(ByRef x As Long, ByRef y As Long)
    Dim tp As tVBBasePoint
    GetCursorPos tp
    x = tp.x
    y = tp.y
End Function

Public Function xGetDispId( _
            ByVal SourceObject As Object, _
            ByVal MethodName As String) _
                As Long
'get method name from IDispatch interface of SourceObject

    Dim oIDispatch As vbBaseTlb.IDispatch
    Dim IID_Null As vbBaseTlb.CLSID

    'get ref to OLE IDispatch interface
    Set oIDispatch = SourceObject

    'get DispatchID for method from IDispatch interface
    '(VB will throw an 'Object Doesn't Support Property Or Method' error on failure)
    oIDispatch.GetIDsOfNames IID_Null, StrConv(MethodName, vbUnicode), 1, 0&, xGetDispId

End Function


Public Function xhWndFromPoint( _
            ByVal x As Long, _
            ByVal y As Long) _
                As Long
    xhWndFromPoint = WindowFromPoint(x, y)
End Function

Public Sub xSleep( _
            ByVal iMilliseconds As Long)
    Sleep iMilliseconds
End Sub

Public Function xSysColor( _
            ByVal iIndex As eSysColor) _
                As Long
    xSysColor = GetSysColor(iIndex)
End Function

Public Function xSystemMetric( _
            ByVal iIndex As eSysMetric) _
                As Long
    xSystemMetric = GetSystemMetrics(iIndex)
End Function
'</Api Utility Interface>

'<Global Object Interface>
Public Function GetGvbBase() _
                As gVbBase
    Set GetGvbBase = Me
End Function
'</Global Object Interface>
'</Public Interface>

Private Sub Class_Initialize()
    InitBitMask
End Sub

Public Sub DisableIDEProtection()
    gbDisableIDEProtection = True
End Sub
