Attribute VB_Name = "Module1"
Option Explicit

Public Const MAX_PATH = 260

Public Type SHFILEINFO
        Hicon As Long                      '  out: icon
        iIcon As Long          '  out: icon index
        dwAttributes As Long               '  out: SFGAO_ flags
        szDisplayName As String * MAX_PATH '  out: display name (or path)
        szTypeName As String * 80         '  out: type name
End Type

Public Const SHGFI_ICON = &H100                         '  get icon
Public Const SHGFI_DISPLAYNAME = &H200                  '  get display name
Public Const SHGFI_TYPENAME = &H400                     '  get type name
Public Const SHGFI_ATTRIBUTES = &H800                   '  get attributes
Public Const SHGFI_ICONLOCATION = &H1000                '  get icon location
Public Const SHGFI_EXETYPE = &H2000                     '  return exe type
Public Const SHGFI_SYSICONINDEX = &H4000                '  get system icon index
Public Const SHGFI_LINKOVERLAY = &H8000                 '  put a link overlay on icon
Public Const SHGFI_SELECTED = &H10000                   '  show icon in selected state
Public Const SHGFI_LARGEICON = &H0                      '  get large icon
Public Const SHGFI_SMALLICON = &H1                      '  get small icon
Public Const SHGFI_OPENICON = &H2                       '  get open icon
Public Const SHGFI_SHELLICONSIZE = &H4                  '  get shell size icon
Public Const SHGFI_PIDL = &H8                           '  pszPath is a pidl
Public Const SHGFI_USEFILEATTRIBUTES = &H10             '  use passed dwFileAttribute

Public Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long

Public Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Public Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
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

Public Const INVALID_HANDLE_VALUE = -1
'Public Const ERROR_NO_MORE_FILES = 18&

Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100
Public Const FILE_ATTRIBUTE_COMPRESSED = &H800
Public Const FILE_ATTRIBUTE_ALL = FILE_ATTRIBUTE_READONLY Or FILE_ATTRIBUTE_HIDDEN _
    Or FILE_ATTRIBUTE_SYSTEM Or FILE_ATTRIBUTE_DIRECTORY Or FILE_ATTRIBUTE_ARCHIVE _
    Or FILE_ATTRIBUTE_NORMAL Or FILE_ATTRIBUTE_TEMPORARY Or FILE_ATTRIBUTE_COMPRESSED
Public Const FILE_ATTRIBUTE_AllButDir = FILE_ATTRIBUTE_ALL And (Not FILE_ATTRIBUTE_DIRECTORY)

Public Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Public Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long

Public Const LVS_SHAREIMAGELISTS = &H40&
Public Const GWL_STYLE = (-16)
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, _
                            ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const LVSIL_NORMAL = 0
Public Const LVSIL_SMALL = 1


Public Const LVM_FIRST = &H1000&
Public Const LVM_SETIMAGELIST = (LVM_FIRST + 3)
Public Const LVM_SETEXTENDEDLISTVIEWSTYLE As Long = (LVM_FIRST + 54)
Public Const LVM_GETEXTENDEDLISTVIEWSTYLE As Long = (LVM_FIRST + 55)

Public Const LVS_EX_FULLROWSELECT As Long = &H20
Public Const LVS_EX_GRIDLINES As Long = &H1
Public Const LVS_EX_CHECKBOXES As Long = &H4
Public Const LVS_EX_HEADERDRAGDROP As Long = &H10
Public Const LVS_EX_TRACKSELECT As Long = &H8

Public Const LVIF_IMAGE = &H2
Public Const LVM_SETITEM = (LVM_FIRST + 6)

Public Type LV_ITEM
    mask As Long
    iItem As Long
    iSubItem As Long
    state As Long
    stateMask As Long
    pszText As String
    cchTextMax As Long
    iImage As Long
    lParam As Long '(~ ItemData)
'#if (_WIN32_IE >= 0x0300)
    iIndent As Long
'#End If
End Type

' =========================================================

Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" _
                            (ByVal lpBuffer As String, _
                            ByVal nSize As Long) As Long

' =========================================================

Declare Sub InitCommonControls Lib "comctl32.dll" ()

' Retrieves the number of images in an image list.
Declare Function ImageList_GetImageCount Lib "comctl32" _
                            (ByVal himl As Long) As Long
                            
' Retrieves the dimensions of images in an image list. All images in an image
' list have the same dimensions. Rtns non-zero if successful, False (0) if failure.
Declare Function ImageList_GetIconSize Lib "comctl32" _
                            (ByVal himl As Long, _
                            cx As Long, _
                            cy As Long) As Boolean

' Creates an icon or cursor based on an image and mask in an image list.
' If the function succeeds, the return value is the handle of the icon or cursor.
' If the function fails, the return value is NULL.
Declare Function ImageList_GetIcon Lib "comctl32" _
                            (ByVal himl As Long, _
                            ByVal I As Long, _
                            ByVal flags As Long) As Long

' i = Index of the image.

' flags = Flag specifying the drawing style. This parameter can be one or more
' of the following values:

' Draws the image using the background color for the image list. If the background
' color is the CLR_NONE value, the image is drawn transparently using the mask.
Public Const ILD_NORMAL = &H0

' Draws the image transparently using the mask, regardless of the background
' color. This value has no effect if the image list does not contain a mask.
Public Const ILD_TRANSPARENT = &H1

' Draws the image, blending 25 percent with the system highlight color. This value
' has no effect if the image list does not contain a mask.
Public Const ILD_BLEND25 = &H2
Public Const ILD_FOCUS = ILD_BLEND25

' Draws the image, blending 50 percent with the system highlight color. This value
' has no effect if the image list does not contain a mask.
Public Const ILD_BLEND50 = &H4
Public Const ILD_SELECTED = ILD_BLEND50
Public Const ILD_BLEND = ILD_BLEND50

' Draws the mask.
Public Const ILD_MASK = &H10

Public Const ILD_IMAGE = &H20

Public Const ILD_OVERLAYMASK = &HF00
' #define INDEXTOOVERLAYMASK(i)   ((i) << 8)

' No background color. The image is drawn transparently.
Public Const CLR_NONE = &HFFFF

' Default background color. The image is drawn using the background color of the
' image list.
Public Const CLR_DEFAULT = &HFF000000
Public Const CLR_HILIGHT = CLR_DEFAULT
 
' =========================================================

Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, _
phiconLarge As Long, phiconSmall As Long, ByVal nIcons As Long) As Long

Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, _
ByVal nIconIndex As Long) As Long

Public Const LR_LOADFROMFILE = &H10

Declare Function DrawIconEx Lib "user32" _
                            (ByVal hDC As Long, _
                             ByVal xLeft As Long, _
                             ByVal yTop As Long, _
                             ByVal Hicon As Long, _
                             ByVal cxWidth As Long, _
                             ByVal cyWidth As Long, _
                             ByVal istepIfAniCur As Long, _
                             ByVal hbrFlickerFreeDraw As Long, _
                             ByVal diFlags As Long) As Boolean

' DrawIconEx() diFlags values:
Public Const DI_MASK = &H1
Public Const DI_IMAGE = &H2
Public Const DI_NORMAL = &H3
Public Const DI_COMPAT = &H4
Public Const DI_DEFAULTSIZE = &H8

Declare Function DestroyIcon Lib "user32" _
                            (ByVal Hicon As Long) As Long














