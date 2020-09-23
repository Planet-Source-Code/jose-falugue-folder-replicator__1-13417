Attribute VB_Name = "FolderReplicatorModule"
Option Explicit
Option Compare Text
Dim Result As Long
Public NroArchivosCopiados As Long, NroArchivosActualizados As Long
Public Const MAX_PATH = 260
'Usado por SystemParametersInfo
Public Const SPI_GETICONMETRICS& = 45
Public Const GByte = 1073741824

Public Const MF_BYPOSITION& = &H400&
Public Const DI_MASK& = 1
Public Const DI_IMAGE& = 2
Public Const DI_NORMAL& = 3
'Usados por Pen Style
Public Const PS_DASH& = 1
Public Const PS_DASHDOT& = 3
Public Const PS_DOT& = 2
Public Const PS_SOLID& = 0
Private Const MAX_COMPUTERNAME_LENGTH = 16


Public Const INVALID_HANDLE_VALUE& = -1, ERROR_NO_MORE_FILES& = 18&

Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type


'Usados por TrackPopUpMenu
Public Const TPM_RIGHTALIGN& = &H8&
Public Const TPM_LEFTBUTTON& = &H0&
 
Public Const BDR_INNER = &HC, BDR_OUTER = &H3
Public Const BDR_RAISED = &H5, BDR_RAISEDINNER = &H4
Public Const BDR_RAISEDOUTER = &H1, BDR_SUNKEN = &HA
Public Const BDR_SUNKENINNER = &H8, BDR_SUNKENOUTER = &H2
Public Const EDGE_BUMP = &H9&, EDGE_ETCHED = &H6&
Public Const EDGE_RAISED = &H5&, EDGE_SUNKEN = &HA&
Public Const BF_ADJUST = &H2000, BF_BOTTOM = &H8
Public Const BF_BOTTOMLEFT = &H9, BF_BOTTOMRIGHT = &HC
Public Const BF_DIAGONAL = &H10, BF_FLAT = &H4000
Public Const BF_LEFT = &H1, BF_MIDDLE = &H800
Public Const BF_MONO = &H8000, BF_RECT = &HF
Public Const BF_RIGHT = &H4, BF_SOFT = &H1000
Public Const BF_TOP = &H2, BF_TOPLEFT = &H3
Public Const BF_TOPRIGHT = &H6
Public Declare Function DrawEdge Lib "user32" (ByVal hdc&, qrc As Rect, ByVal edge&, ByVal grfFlags&) As Long


'Usados por GetSystemMetrics
Public Const SM_CXICON& = 11
Public Const SM_CYICON& = 12
Public Const SM_CXSMICON& = 49
Public Const SM_CYSMICON& = 50
Public Const SPI_GETICONTITLELOGFONT& = 31
Public Const SRCCOPY& = &HCC0020
Public Const SRCAND& = &H8800C6
Public Const SRCINVERT& = &H660046
Public Const SRCPAINT& = &HEE0086
Public Const SRCERASE& = &H440328
Public Const MERGECOPY& = &HC000CA
Public Const MERGEPAINT& = &HBB0226
Public Const DESTORPAT& = &HF3008A
Public Const SRCORPAT& = &HFC008A
Public Const COLOR_BTNFACE& = 15
Public Const COLOR_BTNHIGHLIGHT& = 20
Public Const COLOR_BTNHILIGHT& = COLOR_BTNHIGHLIGHT
Public Const COLOR_BTNSHADOW& = 16
Public Const COLOR_BTNTEXT& = 18
Public Const MF_BITMAP& = &H4&

Public Const OPAQUE& = 2
Public Const TRANSPARENT& = 1

Public Const PATCOPY& = &HF00021
Public Const PATINVERT& = &H5A0049
Public Const PATPAINT& = &HFB0A09
Public Const WHITENESS = &HFF0062

Public Const COLOR_3DDKSHADOW& = 21
Public Const COLOR_3DFACE& = COLOR_BTNFACE
Public Const COLOR_3DHIGHLIGHT& = COLOR_BTNHIGHLIGHT
Public Const COLOR_3DHILIGHT& = COLOR_BTNHIGHLIGHT
Public Const COLOR_3DLIGHT& = 22
Public Const COLOR_3DSHADOW& = COLOR_BTNSHADOW
Public Const COLOR_ACTIVEBORDER& = 10
Public Const COLOR_ACTIVECAPTION& = 2
Public Const COLOR_ADJ_MAX& = 100
Public Const COLOR_ADJ_MIN& = -100
Public Const COLOR_APPWORKSPACE& = 12
Public Const COLOR_BACKGROUND& = 1
Public Const COLOR_CAPTIONTEXT& = 9
Public Const COLOR_DESKTOP& = COLOR_BACKGROUND
Public Const COLOR_GRAYTEXT& = 17
Public Const COLOR_HIGHLIGHT& = 13
Public Const COLOR_HIGHLIGHTTEXT& = 14
Public Const COLOR_INACTIVEBORDER& = 11
Public Const COLOR_INACTIVECAPTION& = 3
Public Const COLOR_INACTIVECAPTIONTEXT& = 19
Public Const COLOR_INFOBK& = 24
Public Const COLOR_INFOTEXT& = 23
Public Const COLOR_MENU& = 4
Public Const COLOR_MENUTEXT& = 7
Public Const COLOR_SCROLLBAR& = 0
Public Const COLOR_WINDOW& = 5
Public Const COLOR_WINDOWFRAME& = 6
Public Const COLOR_WINDOWTEXT& = 8




'Usados por DrawText
Public Const DT_BOTTOM& = &H8
Public Const DT_CALCRECT& = &H400
Public Const DT_CENTER& = &H1
Public Const DT_CHARSTREAM& = 4
Public Const DT_DISPFILE& = 6
Public Const DT_EDITCONTROL& = &H2000
Public Const DT_END_ELLIPSIS& = &H8000
Public Const DT_EXPANDTABS& = &H40
Public Const DT_EXTERNALLEADING& = &H200
Public Const DT_INTERNAL& = &H1000
Public Const DT_LEFT& = &H0
Public Const DT_METAFILE& = 5
Public Const DT_MODIFYSTRING& = &H10000
Public Const DT_NOCLIP& = &H100
Public Const DT_NOPREFIX& = &H800
Public Const DT_PATH_ELLIPSIS& = &H4000
Public Const DT_PLOTTER& = 0
Public Const DT_RASCAMERA& = 3
Public Const DT_RASDISPLAY& = 1
Public Const DT_RASPRINTER& = 2
Public Const DT_RIGHT& = &H2
Public Const DT_RTLREADING& = &H20000
Public Const DT_SINGLELINE& = &H20
Public Const DT_TABSTOP& = &H80
Public Const DT_TOP& = &H0
Public Const DT_VCENTER& = &H4
Public Const DT_WORD_ELLIPSIS& = &H40000
Public Const DT_WORDBREAK& = &H10
Public Const LF_FACESIZE& = 32
Public Const LF_FULLFACESIZE& = 64

Public Const WM_LBUTTONDBLCLK& = &H203
Public Const WM_LBUTTONDOWN& = &H201
Public Const WM_LBUTTONUP& = &H202
Public Const WM_MOUSEMOVE& = &H200
 
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2

Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4

Public Const HKEY_CURRENT_USER = &H80000001
Public Const KEY_QUERY_VALUE& = &H1
Public Const KEY_SET_VALUE& = &H2
Public Const KEY_CREATE_SUB_KEY& = &H4
Public Const STANDARD_RIGHTS_WRITE& = &H20000
 
Public Type POINT_Integer
    x As Integer
    y As Integer
End Type

Public Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

 ' // Tray notification definitions
Type NotifyIconData
        cbSize As Long
        hwnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type

Public Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    'lfFacename As String * LF_FACESIZE
   lfFaceName(LF_FACESIZE - 1) As Byte
End Type

Public Type SIZE
    cx As Long
    cy As Long
End Type

Public Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Type ICONMETRICS
    cbSize As Long
    iHorzSpacing As Long
    iVertSpacing As Long
    iTitleWrap As Long
    lfFont As LOGFONT
End Type


Public Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type
Type BROWSEINFO 'Usado por ShBrowseForFolder
    hwndOwner As Long 'hWnd de la ventana Padre del cuadro que se abre
    pidlRoot As Long  'Indica cual será la carpeta raíz , si es cero es el escritorio
    pszDisplayName As String 'Buffer donde se guarda la carpeta elegida
    lpszTitle As String 'Mensaje que aparece en el cuadro
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Public Type SHFILEOPSTRUCT
        hwnd As Long
        wFunc As Long
        pFrom As String
        pTo As String
        fFlags As Long
        fAnyOperationsAborted As Long
        hNameMappings As Long
        lpszProgressTitle As String '  only used if FOF_SIMPLEPROGRESS
End Type
'Flags usados por SHFileOperation
Public Const FOF_ALLOWUNDO = &H40
Public Const FOF_CONFIRMMOUSE = &H2
Public Const FOF_FILESONLY = &H80
Public Const FOF_MULTIDESTFILES = &H1
Public Const FOF_NOCONFIRMATION = &H10
Public Const FOF_NOCONFIRMMKDIR = &H200
Public Const FOF_RENAMEONCOLLISION = &H8
Public Const FOF_SILENT = &H4
Public Const FOF_SIMPLEPROGRESS = &H100
Public Const FOF_WANTMAPPINGHANDLE = &H20
Public Const FOF_CREATEPROGRESSDLG = &H0
'Función ejecutada por SHFileOperation
Public Const FO_COPY = &H2
Public Const FO_DELETE = &H3
Public Const FO_MOVE = &H1
Public Const FO_RENAME = &H4
Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As Any) As Long

'ulFlags puede tener los siguientes valores
Public Const BIF_BROWSEFORCOMPUTER = 0, BIF_BROWSEFORPRINTER = 1
Public Const BIF_DONTGOBELOWDOMAIN = 2, BIF_RETURNFSANCESTORS = 3
Public Const BIF_RETURNONLYFSDIRS = 4, BIF_STATUSTEXT = 5

Declare Function SHBrowseForFolder& Lib "shell32" Alias "SHBrowseForFolderA" (Param1 As BROWSEINFO)
Declare Function GetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal IDList As Long, ByVal Buffer As String) As Long


'Valores que toma nFolder en SHGetSpecialFolderLocation
Public Const CSIDL_DESKTOPDIRECTORY = 0& ' Escritorio del sistema de archivos donde se colocan los objetos
Public Const CSIDL_PROGRAMS = 2&  ' Grupos de programas
Public Const CSIDL_CONTROLS = 3 'Carpeta virtual que contiene íconos de aplicac. del Panel de Control
Public Const CSIDL_PRINTERS = 4 'Carpeta Virtual de impresoras instaladas
Public Const CSIDL_PERSONAL = 5 ' Repositorio para documentos
Public Const CSIDL_TEMPLATES = 6  ' Donde se guardan las plantillas
Public Const CSIDL_STARTUP = 7         ' Menú Inicio
Public Const CSIDL_RECENT = 8      ' Documentos mas recientes
Public Const CSIDL_SENDTO = 9      'Items de menus de Enviar a:
Public Const CSIDL_BITBUCKET = 10 ' Devuelve la Papelera
Public Const CSIDL_STARTMENU = 11  'Items del Menú Inicio Principal
Public Const CSIDL_DRIVES = 17 ' Carpeta Virtual - Mi PC
Public Const CSIDL_FONTS = 20  'Carpeta Virtual que contiene las fuentes
Public Const CSIDL_SHELLNEW = 21
Public Const CSIDL_DESKTOP = 2   'Escritorio de Windows Carpeta Virtual en root de NameSpace
Public Const CSIDL_NETHOOD = 6 ' Carpeta Virtual que contiene objetos que aparecen el Network Neighborhood
Public Const CSIDL_NETWORK = 7 'Carpeta Virtual que representa el tope del Network

'' Constantes que usa SHGetFileInfo========================================
Public Const SHGFI_DISPLAYNAME& = &H200
Public Const SHGFI_EXETYPE& = &H2000
Public Const SHGFI_ICON& = &H100
Public Const SHGFI_ICONLOCATION& = &H1000
Public Const SHGFI_LARGEICON& = &H0
Public Const SHGFI_LINKOVERLAY& = &H8000
Public Const SHGFI_OPENICON& = &H2
Public Const SHGFI_PIDL& = &H8
Public Const SHGFI_SELECTED& = &H10000
Public Const SHGFI_SHELLICONSIZE& = &H4
Public Const SHGFI_SMALLICON& = &H1
Public Const SHGFI_SYSICONINDEX& = &H4000
Public Const SHGFI_TYPENAME& = &H400
Public Const SHGFI_USEFILEATTRIBUTES& = &H10
'Usados por SendMessage para seleccionar un item en listbox
Public Const LB_SETSEL& = &H185
Public Const LB_SELECTSTRING& = &H18C


Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Public Declare Function FindFirstFile& Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA)
Public Declare Function FindNextFile& Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA)
Public Declare Function FindClose& Lib "kernel32" (ByVal hFindFile As Long)
Public Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long

Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Public Const DRIVE_REMOVABLE = 2, DRIVE_FIXED = 3, DRIVE_CDROM = 5
Public Const DRIVE_REMOTE = 4, DRIVE_RAMDISK = 6



Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd&, ByVal wMsg&, ByVal wParam&, ByVal lParam As Any) As Long
Public Declare Function SHGetSpecialFolderLocation& Lib "shell32" (ByVal hwndOwner&, ByVal nFolder&, ppidl&)
Public Declare Function GetSystemDirectory& Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long)
Public Declare Function GetWindowsDirectory& Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long)
Public Declare Function Shell_NotifyIcon& Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage&, lpData As NotifyIconData)
Public Declare Function DrawAnimatedRects& Lib "user32" (ByVal hwnd As Long, ByVal idAni As Long, lprcFrom As Rect, lprcTo As Rect)
Public Declare Function GetDC& Lib "user32" (ByVal hwnd As Long)
Public Declare Function ReleaseDC& Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long)

' Private Const KEY_WRITE& = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (NotSYNCHRONIZE))
Declare Function RegCloseKey& Lib "advapi32.dll" (ByVal hKey&)
Declare Function RegDeleteKey& Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey&, ByVal lpSubKey As String)
Declare Function RegOpenKeyEx& Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey&, ByVal lpSubKey As String, ByVal ulOptions&, ByVal samDesired&, phkResult&)


'Para acceder al archivo pahts.ini
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long

'Abre el Cuadro de diálogo de Formatear Drive
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey&) As Long
Declare Function ShellAbout& Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd&, ByVal szApp$, ByVal szOtherStuff$, ByVal hIcon&)

Declare Function SetRect Lib "user32" (lpRect As Rect, ByVal X1&, ByVal Y1&, ByVal X2&, ByVal Y2&) As Long
Declare Function LockWindowUpdate& Lib "user32" (ByVal hwndLock&)
Declare Function ShellExecute& Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd&, ByVal lpOperation$, ByVal lpFile$, ByVal lpParameters$, ByVal lpDirectory$, ByVal nShowCmd&)
Declare Function ExtractAssociatedIcon& Lib "shell32.dll" Alias "ExtractAssociatedIconA" (ByVal hInst&, ByVal lpIconPath As String, lpiIcon&)
Declare Function DrawFocusRect& Lib "user32" (ByVal hdc As Long, lpRect As Rect)
Declare Function SetCapture& Lib "user32" (ByVal hwnd&)
Declare Function SHGetFileInfoByIDList& Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As Long, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long)
Declare Function SHGetFileInfo& Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long)
Declare Function ReleaseCapture& Lib "user32" ()
Declare Function CreatePen& Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long)
Declare Function GetCursorPos& Lib "user32" (lpPoint As POINTAPI)
Declare Function DrawIcon& Lib "user32" (ByVal hdc&, ByVal x&, ByVal y&, ByVal hIcon&)
Declare Function LoadIconByNum& Lib "user32" Alias "LoadIconA" (ByVal hInstance&, ByVal lpIconName&)
Declare Function DestroyIcon& Lib "user32" (ByVal hIcon&)
Declare Function ExtractIcon& Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst&, ByVal lpszExeFileName As String, ByVal nIconIndex&)
Declare Function ExtractIconEx& Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex&, phIconLarge&, phIconSmall&, ByVal nIcons&)
Declare Function DrawIconEx& Lib "user32" (ByVal hdc&, ByVal xLeft&, ByVal yTop&, ByVal hIcon&, ByVal cxWidth&, ByVal cyWidth&, ByVal istepIfAniCur&, ByVal hbrFlickerFreeDraw&, ByVal diFlags&)
Declare Function FindExecutable& Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String)
Declare Function DrawText& Lib "user32" Alias "DrawTextA" (ByVal hdc&, ByVal lpStr As String, ByVal nCount&, lpRect As Rect, ByVal wFormat&)
Declare Function GetDesktopWindow& Lib "user32" ()
Declare Function TrackPopupMenu& Lib "user32" (ByVal hMenu&, ByVal wFlags&, ByVal x&, ByVal y&, ByVal nReserved&, ByVal hwnd&, lprc As Rect)
Declare Function TrackPopupMenuByNum& Lib "user32" Alias "TrackPopupMenu" (ByVal hMenu&, ByVal wFlags&, ByVal x&, ByVal y&, ByVal nReserved&, ByVal hwnd&, ByVal lprc&)
Declare Function GetTickCount Lib "kernel32" () As Long

'Public Declare Function TrackPopupMenuEx& Lib "user32" (ByVal hMenu&, ByVal un&, ByVal n1&, ByVal n2&, ByVal hWnd&, lpTPMParams As TPMPARAMS)

Declare Function GetTextExtentPoint32& Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc&, ByVal lpsz As String, ByVal cbString&, lpSize As SIZE)

Declare Function SystemParametersInfo& Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction&, ByVal uParam&, lpvParam As Any, ByVal fuWinIni&)
Declare Function GetSystemMetrics& Lib "user32" (ByVal nIndex&)
Declare Function CreateFontIndirect& Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT)

Declare Function GetMenu& Lib "user32" (ByVal hwnd&)
Declare Function GetMenuCheckMarkDimensions Lib "user32" () As POINT_Integer
Declare Function GetSubMenu& Lib "user32" (ByVal hMenu&, ByVal nPos&)
Declare Function SetMenu& Lib "user32" (ByVal hwnd&, ByVal hMenu&)
Declare Function SetMenuItemBitmaps& Lib "user32" (ByVal hMenu&, ByVal nPosition&, ByVal wFlags&, ByVal hBitmapUnchecked&, ByVal hBitmapChecked&)
Declare Function SelectObject& Lib "gdi32" (ByVal hdc&, ByVal hObject&)
Declare Function DeleteObject& Lib "gdi32" (ByVal hObject&)
Declare Function GetObjectAPI& Lib "gdi32" Alias "GetObjectA" (ByVal hObject&, ByVal nCount&, lpObject As Any)
Declare Function DeleteDC& Lib "gdi32" (ByVal hdc&)
Declare Function CreateCompatibleBitmap& Lib "gdi32" (ByVal hdc&, ByVal nWidth&, ByVal nHeight&)
Declare Function CreateCompatibleDC& Lib "gdi32" (ByVal hdc&)
Declare Function Rectangle& Lib "gdi32" (ByVal hdc&, ByVal X1&, ByVal Y1&, ByVal X2&, ByVal Y2&)

Declare Function SetStretchBltMode& Lib "gdi32" (ByVal hdc&, ByVal nStretchMode&)
Declare Function StretchBlt& Lib "gdi32" (ByVal hdc&, ByVal x&, ByVal y&, ByVal nWidth&, ByVal nHeight&, ByVal hSrcDC&, ByVal xSrc&, ByVal ySrc&, ByVal nSrcWidth&, ByVal nSrcHeight&, ByVal dwRop&)
Declare Function CreateSolidBrush& Lib "gdi32" (ByVal crColor&)
Declare Function SetBkMode& Lib "gdi32" (ByVal hdc&, ByVal nBkMode&)
Declare Function ModifyMenuBynum& Lib "user32" Alias "ModifyMenuA" (ByVal hMenu&, ByVal nPosition&, ByVal wFlags&, ByVal wIDNewItem&, ByVal lpString&)
Declare Function GetMenuItemID& Lib "user32" (ByVal hMenu&, ByVal nPos&)
Declare Function SetBkColor& Lib "gdi32" (ByVal hdc&, ByVal crColor&)
Declare Function PatBlt& Lib "gdi32" (ByVal hdc&, ByVal x&, ByVal y&, ByVal nWidth&, ByVal nHeight&, ByVal dwRop&)
Declare Function GetNearestColor& Lib "gdi32" (ByVal hdc&, ByVal crColor&)
Declare Function GetSysColor& Lib "user32" (ByVal nIndex&)
Declare Function BitBlt& Lib "gdi32" (ByVal hDestDC&, ByVal x&, ByVal y&, ByVal nWidth&, ByVal nHeight&, ByVal hSrcDC&, ByVal xSrc&, ByVal ySrc&, ByVal dwRop&)
Declare Function CreateBitmapIndirect& Lib "gdi32" (lpBitmap As BITMAP)
Declare Function GetObjectType& Lib "gdi32" (ByVal hgdiobj&)

' Abre un cuadro en el que se elige el programa para abrir el documento que le paso en NameFile$
Declare Function OpenAs_RunDLL& Lib "shell32" (ByVal Flag1&, ByVal Flag2&, ByVal NameFile$, ByVal Flag3&)
Public Declare Function DrawFrameControl& Lib "user32" (ByVal hdc As Long, lpRect As Rect, ByVal un1 As Long, ByVal un2 As Long)
Public Const DFC_BUTTON& = 4
Public Const DFC_CAPTION& = 1
Public Const DFC_MENU& = 2
Public Const DFC_SCROLL& = 3
Public Const DFCS_ADJUSTRECT& = &H2000
Public Const DFCS_BUTTON3STATE& = &H8
Public Const DFCS_BUTTONCHECK& = &H0
Public Const DFCS_BUTTONPUSH& = &H10
Public Const DFCS_BUTTONRADIO& = &H4
Public Const DFCS_BUTTONRADIOIMAGE& = &H1
Public Const DFCS_BUTTONRADIOMASK& = &H2
Public Const DFCS_CAPTIONCLOSE& = &H0
Public Const DFCS_CAPTIONHELP& = &H4
Public Const DFCS_CAPTIONMAX& = &H2
Public Const DFCS_CAPTIONMIN& = &H1
Public Const DFCS_CAPTIONRESTORE& = &H3
Public Const DFCS_CHECKED& = &H400
Public Const DFCS_FLAT& = &H4000
Public Const DFCS_INACTIVE& = &H100
Public Const DFCS_MENUARROW& = &H0
Public Const DFCS_MENUARROWRIGHT& = &H4
Public Const DFCS_MENUBULLET& = &H2
Public Const DFCS_MENUCHECK& = &H1
Public Const DFCS_MONO& = &H8000
Public Const DFCS_PUSHED& = &H200
Public Const DFCS_SCROLLCOMBOBOX& = &H5
Public Const DFCS_SCROLLDOWN& = &H1
Public Const DFCS_SCROLLLEFT& = &H2
Public Const DFCS_SCROLLRIGHT& = &H3
Public Const DFCS_SCROLLSIZEGRIP& = &H8
Public Const DFCS_SCROLLSIZEGRIPRIGHT& = &H10
Public Const DFCS_SCROLLUP& = &H0




Public Const MAX_ENT_MENU = 16

Public Type DATOSEXE
    Path As String
    DisplayName As String
    Param As String
    ShowWindow As Long 'Forma de abrir la ventana al inicio
    hIcon As Long 'handle al icono que muestra en la barra
    hBitmap As Long 'handle al bitmap del menu
    NroDeIcono As Long
    IconPath As String
End Type
   



    






Public Function Tray(ByVal Msg As Long, ByVal flag As Long, ByVal Tip As String, ByVal hIcon1 As Long)

Dim NotifIcon As NotifyIconData
' Dibuja un ícono en la bandeja de entrada
' en Flag debo pasarle NIF_ICON para agregarle un ícono,NIF_TIP para el tooltip etc.
' en Msg que quiero hacer agregar el ícono,modificarle algo o borrarlo
    
With NotifIcon
    .cbSize = Len(NotifIcon)
    .hwnd = Screen.ActiveForm.hwnd
    .uID = Screen.ActiveForm.hwnd
     .uFlags = flag
     .hIcon = hIcon1
     .uCallbackMessage = WM_LBUTTONDOWN
     .szTip = Tip + vbNullChar
    
End With


Tray = Shell_NotifyIcon(Msg, NotifIcon)

End Function

Function GetMouseButtons() As Long
Dim LBot As Long, RBot As Long, MBot As Long
'Devuelve el estado de los botones del mouse
'Devuelve  los siguientes valores:
'Botón Izq. apretado = 1
'Botón Derecho apretado = 2
'Los Dos Botones apretados = 3
' GetAsynKeyState pone el bit 15 en 1 si la tecla está siendo apretada en el momento
'que llamo a la función y el bit 0 en 1 si la tecla fue apretada la última vez que llamé
' a GetAsyncKeyState

LBot = (GetAsyncKeyState(vbLeftButton) And &H8000) / &H8000
RBot = (GetAsyncKeyState(vbRightButton) And &H8000) / &H8000
MBot = (GetAsyncKeyState(vbMiddleButton) And &H8000) / &H8000

GetMouseButtons = LBot * vbLeftButton + RBot * vbRightButton + MBot * vbMiddleButton

End Function




Public Sub About(Formu As Form)
Dim szApp As String, szOtherStuff As String
szApp = "Folder Replicator  - AguSoft Corporation"
szOtherStuff = "Author: © José Luis Falugue" & vbCrLf & "ALL RIGHTS RESERVED" & vbCrLf
ShellAbout Formu.hwnd, szApp, szOtherStuff, Formu.Icon


End Sub


Function GetStrNro(ByVal Cad As String, ByVal Nro As Integer, Optional Separador) As String
'Cad= Cadena con items separados por comas
'Nro= Nro de item que devuelve GetStrNro
'Separador de un ítem del otro , si no lo paso,asumo que es una coma
Dim Pos As Integer, Aux As String, x As Integer
If IsMissing(Separador) Then
    Separador = ","
End If
Cad = IIf(Right$(Cad, 1) <> Separador, Cad & Separador, Cad)

For x = 1 To Nro
    Pos = InStr(Cad, Separador)
        If Pos = 0 Then
                Aux = ""
                Exit For
        End If
    Aux = Left$(Cad, Pos - 1)
    Cad = Mid$(Cad, Pos + 1)
Next x
GetStrNro = Aux
End Function


Public Function RegEraseKey(ByVal Key As String) As Long
Const VBKey = "Software\VB and VBA Program Settings"
Dim Result As Long
Dim hKey As Long
Result = RegOpenKeyEx(HKEY_CURRENT_USER, VBKey, 0, 0, hKey)
Result = RegDeleteKey(hKey, Key)
Result = RegCloseKey(hKey)
End Function

Function GetIconMetrics() As ICONMETRICS
Dim IcnMetr As ICONMETRICS, Result As Long
IcnMetr.cbSize = Len(IcnMetr)
Result = SystemParametersInfo(SPI_GETICONMETRICS, 0, IcnMetr, 0)
GetIconMetrics = IcnMetr

End Function

Public Function jlfExtractAssociatedIcon(ByVal FilePath, ByVal NroDeIcono As Long) As Long
' Dim NroResource As Long
FilePath = FilePath + vbNullChar + Space$(MAX_PATH)
Result = ExtractAssociatedIcon(App.hInstance, FilePath, NroDeIcono)
'FilePath = Trim$(FilePath)
jlfExtractAssociatedIcon = Result
End Function

Public Function CenterForm(MyForm As Form)
MyForm.Left = (Screen.Width - MyForm.Width) / 2
MyForm.Top = (Screen.Height - MyForm.Height) / 2
End Function
Public Sub jlfDrawAnimatedRects(MyForm As Form)
Dim RectIni As Rect, RectFinal As Rect
Dim DesktopDC As Long
Dim TwipsX As Long, TwipsY As Long
Dim Pen As Long, OldPen As Long
TwipsX = Screen.TwipsPerPixelX
TwipsY = Screen.TwipsPerPixelY
DesktopDC = GetDC(0)
Pen = CreatePen(PS_DOT, 5, vbBlack)
OldPen = SelectObject(DesktopDC, Pen)
SetRect RectIni, Screen.Width \ TwipsX, Screen.Height \ TwipsY, Screen.Width \ TwipsX - 1, Screen.Height \ TwipsY - 1
With MyForm
    SetRect RectFinal, .Left \ TwipsX, .Top \ TwipsY, (.Left + .Width) \ TwipsX, (.Top + .Height) \ TwipsY
End With
Result = DrawAnimatedRects(0, 0, RectIni, RectFinal)

Result = SelectObject(DesktopDC, OldPen)
Result = ReleaseDC(0, DesktopDC)
Result = DeleteObject(Pen)
End Sub

Function GetPath(ByVal FilePath As String) As String
Dim x As Integer, Aux As String * 1
For x = Len(FilePath) To 1 Step -1
    FilePath = Left$(FilePath, x)
    Aux = Right$(FilePath, 1)
    If Aux = "\" Or Aux = ":" Then Exit For
Next x
GetPath = FilePath

End Function

Public Function BuscarCarpeta(ByVal FolderID As Long, ByVal Flags As Long) As String
' Abre una ventana que me permite buscar por todas las carpetas
' BuscarCarpeta devuelve el path completa de la carpeta elegida
'en FolderID puedo pasarle el Número que identifica a las carpetas del sistema
'Están guardadas en las constantes que empiezan con CSIDL_****
'Si le paso como FolderID= 1 muestra la carpeta completa
Dim DatosArch As BROWSEINFO
Dim PathStr As String
Dim Result As Long
Dim IDList As Long
Dim Buffer As String

Buffer = Space$(MAX_PATH)
IDList = GetSpecialFolderID(FolderID)
With DatosArch
    .hwndOwner = Screen.ActiveForm.hwnd
    .pidlRoot = IDList
    .pszDisplayName = Buffer
    .lpszTitle = "Choose a folder to add"
    .ulFlags = Flags
    .lpfn = 0
    .lParam = 0
    .iImage = 0
End With
' SHBrowseForFolder devuelve un número que identifica a cada carpeta
IDList = SHBrowseForFolder(DatosArch)
PathStr = Space$(MAX_PATH)

Result = GetPathFromIDList(IDList, PathStr)
PathStr = Trim$(PathStr)
If IDList = 0 Then
    BuscarCarpeta = vbNullString
ElseIf PathStr = vbNullChar Then ' GetPathFromIDList no me devolvió ningún Path en PathStr
    BuscarCarpeta = GetDisplayName(IDList) ' Obtengo el nombre del sistema
Else
    BuscarCarpeta = Left$(PathStr, InStr(PathStr, vbNullChar) - 1) 'Muestro el Path de la carpeta que elegí
End If
End Function
Public Function GetSpecialFolderID(ByVal FolderID As Long) As Long
Dim Result As Long, IDList As Long
Result = SHGetSpecialFolderLocation(0, FolderID, IDList)
GetSpecialFolderID = IDList
End Function
Public Function GetSpecialFolderName(ByVal IDFolder As Long) As String
Dim Buffer As String
Dim Result As Long, IDList As Long
Buffer = Space(MAX_PATH)
Result = SHGetSpecialFolderLocation(0, IDFolder, IDList)
If Result >= 0 Then
    ' Debug.Print Result, IDList
    Result = GetPathFromIDList(IDList, Buffer)
    ' Debug.Print "Display Name>> " & GetDisplayName(IDList)
End If
 GetSpecialFolderName = Buffer
End Function

Public Function GetDisplayName(ByVal IDList As Long) As String
'SHGetFileInfo& (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long)
Dim FileInf As SHFILEINFO
Dim Result As Long
' FileInf.dwAttributes = vbArchive
Result = SHGetFileInfoByIDList(IDList, 0, FileInf, Len(FileInf), SHGFI_PIDL + SHGFI_DISPLAYNAME)
GetDisplayName = FileInf.szDisplayName
End Function
Function GetBaseName(ByVal Source As String) As String
' Devuelve el nombre y extensión del archivo
' No verifica que exista
Do While InStr(Source, "\") <> 0
    Source = Mid$(Source, InStr(Source, "\") + 1)
Loop
If InStr(Source, ":") <> 0 Then
    Source = Mid$(Source, InStr(Source, ":") + 1)
End If
GetBaseName = Source
End Function


Public Function CompletarPath(ByVal PathOrig As String) As String
If PathOrig = "" Then
    CompletarPath = ""
    Exit Function
End If
CompletarPath = IIf(Right$(PathOrig, 1) <> "\", PathOrig & "\", PathOrig)
End Function
Public Function FileExist(ByVal PathCompleto As String) As Boolean
'Puedo pasarle una carpeta o un archivo en Objeto
Const Todos = vbDirectory Or vbHidden Or vbSystem
On Error Resume Next
FileExist = Dir(PathCompleto, Todos) <> ""
If Err Then
    MsgBox Err.Description, vbCritical Or vbRetryCancel
End If
    

End Function

Public Function ShellCopyFile(ByVal FilesToCopy As String, ByVal PathDest As String)
Dim FileOp As SHFILEOPSTRUCT
'Dim LenFileOp As Long
'Dim FileOpBuffer() As Byte
' The SHFILOPSTRUCT is not double word aligned.  If no steps are
' taken, the last 3 variables will not be passed correctly.  This
' has no impact UNLESS THE PROGRESS TITLE NEED TO BE CHANGED!!!

'LenFileOp = LenB(FileOp)    ' double word alignment increase the
'ReDim FileOpBuffer(1 To LenFileOp) ' size of the structure.
If Not FileExist(PathDest) Then
    Result = MakeAllDir(PathDest)
    If Result <> 0 Then
        ShellCopyFile = Result
    End If
End If
With FileOp
    .hwnd = FolderReplicatorFrm.hwnd
    .wFunc = FO_COPY
    .pFrom = FilesToCopy & vbNullChar
    .pTo = PathDest
    .fFlags = FOF_FILESONLY Or FOF_NOCONFIRMATION 'Or FOF_SIMPLEPROGRESS
    '.lpszProgressTitle = AcortarString("Copying File: " & PathDest, 50)
End With

'Copiamos la estructura en un byte array
'Call CopyMemory(FileOpBuffer(1), FileOp, LenFileOp)

'Ahora movemos los últimos 12 bytes 2 lugares para alinear los datos
'Call CopyMemory(FileOpBuffer(19), FileOpBuffer(21), 12)
Result = SHFileOperation(FileOp) 'Buffer(1))
ShellCopyFile = Result
DoEvents
End Function
Public Function GetSizeOfPath(ByVal Path As String) As Long
'Devuelve el tamaño del path que le paso como parámetro en Path
'Si le paso C:\ por ejemplo devuelve el tamaño real ocupado por todos los archivos y directorios
'en el disco C: , el tamaño de cada archivo lo obtengo en la estructura FileData y cada subdirectorio
'guarda una entrada de directorio por cada archivo o subdirectorio que contenga.

Dim BranchSize As Long
Dim hFindFile As Long, Sigo As Long
Dim FileData As WIN32_FIND_DATA
Dim FileName As String
'Dim EntDir As Long 'Número de entradas de directorio que ocupa cada nombre largo
Path = IIf(Right$(Path, 1) <> "\", Path & "\", Path)
hFindFile = FindFirstFile(Path & "*.*", FileData)
' Si a hFindFile le paso un directorio raíz vacío p.ej. A:\'devuelve INVALID_HANDLE_VALUE.
' Entonces no me sirve para calcular el tamaño de la rama.
' Uso entonces GetDriveType(). Si el drive es válido me devuelve
' DRIVE_REMOVABLE = 2,DRIVE_FIXED = 3,DRIVE_REMOTE = 4, DRIVE_CDROM = 5
' DRIVE_RAMDISK = 6. Si devuelve 0 el drive no pudo ser identificado, si
' devuelve 1 el Drive no existe.

If hFindFile = INVALID_HANDLE_VALUE Then
    If GetDriveType(Path) = 1 Then 'El drive no existe
            GetSizeOfPath = -1
    Else
            GetSizeOfPath = 0
    End If
    Exit Function
End If
With FileData
Do
    '*** Si es un subdirectorio sumo el tamaño de esa carpeta
    If (.dwFileAttributes And vbDirectory) Then
        If .cFileName <> "." And .cFileName <> ".." Then  'Calculo el tamaño de la Rama
            FileName = Left$(.cFileName, InStr(.cFileName, vbNullChar) - 1)
            BranchSize = BranchSize + GetSizeOfPath(Path & FileName)
        End If
    Else 'Si no es directorio sumo el tamaño del archivo
        BranchSize = BranchSize + .nFileSizeLow
    End If
    
    Sigo = FindNextFile(hFindFile, FileData)
Loop While Sigo <> 0
End With
'Calculo los clusters que ocupan los Subdirectorios
FindClose hFindFile
GetSizeOfPath = BranchSize

End Function

Public Function AcortarString(ByVal Cad As String, MaxLen As Long) As String
'Esta función acorta Cad si el largo es superior a MaxLen
'Si el largo de Cad es Menor o igual a MaxLen devuelve Cad.
Dim AuxCad As String, Largo As Long
Largo = (MaxLen - 3) \ 2
AcortarString = Cad
If Len(Cad) > MaxLen Then
    AuxCad = Left$(Cad, Largo) & "..." & Right$(Cad, Largo)
    AcortarString = AuxCad
End If
    

End Function


Public Function MakeAllDir(ByVal Path As String) As Long
'Esta función construye todo el path que se le pasa en Path
'P.ej: Si le paso "C:\temp\dir1\dir2" construye temp luego dir1 y luego dir2
'Si encuentra algún error devuelve el código de error si es OK devuelve cero.
Dim Folder As String, Posic As Long
Dim OldPosic As Long
OldPosic = 1
Path = CompletarPath(Path)
MakeAllDir = 0
Do
    On Error Resume Next
    Posic = InStr(OldPosic + 1, Path, "\")
    Folder = Left$(Path, Posic)
    If Not FileExist(Folder) Then
        'Debug.Print Folder
        MkDir Folder
        If Err Then MakeAllDir = Err.Number
    End If
    OldPosic = Posic
Loop While Posic < Len(Path)

End Function


Public Sub SortStringArray(Vector() As String) ', Ultimo As Long)
'Ordena un array de strings
'el indice del array debe comenzar en 0, Array(0 to n)
Dim PuntIzq As Long, PuntFijo As Long, PuntDer As Long
Dim Delta As Long, LargoDer As Long
Dim Ultimo As Long
Dim TempStr As String

Ultimo = UBound(Vector) + 1
Delta = Ultimo \ 2

Do Until Delta = 0
    PuntFijo = 0
    PuntIzq = 0
    LargoDer = Ultimo - Delta

    Do While PuntFijo < LargoDer
        PuntDer = PuntIzq + Delta
        
        If Vector(PuntIzq) > Vector(PuntDer) Then 'Swap
            TempStr = Vector(PuntIzq)
            Vector(PuntIzq) = Vector(PuntDer)
            Vector(PuntDer) = TempStr
            
            If PuntIzq >= Delta Then
                PuntIzq = PuntIzq - Delta
            End If
        Else
            PuntFijo = PuntFijo + 1
            PuntIzq = PuntFijo
        End If
    Loop
    Delta = Delta \ 2
Loop
                             
                               
End Sub
Public Sub prueba()
Dim x As Long
Dim a() As String
ReDim a(0)
a(0) = 1
Debug.Print UBound(a)
End Sub
