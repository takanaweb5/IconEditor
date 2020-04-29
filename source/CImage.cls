VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As LongPtr
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Enum GpStatus
   Ok = 0
   GenericError = 1
   InvalidParameter = 2
   OutOfMemory = 3
   ObjectBusy = 4
   InsufficientBuffer = 5
   NotImplemented = 6
   Win32Error = 7
   WrongState = 8
   Aborted = 9
   FileNotFound = 10
   ValueOverflow = 11
   AccessDenied = 12
   UnknownImageFormat = 13
   FontFamilyNotFound = 14
   FontStyleNotFound = 15
   NotTrueTypeFont = 16
   UnsupportedGdiplusVersion = 17
   GdiplusNotInitialized = 18
   PropertyNotFound = 19
   PropertyNotSupported = 20
End Enum

Private Type Guid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Private Type PICTDESC_ICON
    Size   As Long
    Type   As Long
    hIcon  As LongPtr
End Type

Private Type PICTDESC_BMP
    Size    As Long
    Type    As Long
    hBitmap As LongPtr
    hPal    As LongPtr
End Type

Private Type ICONDIR
    Reserved As Integer
    Type     As Integer
    Count    As Integer
End Type

Private Type ICONDIRENTRY
    Width      As Byte
    Height     As Byte
    ColorCount As Integer
    Reserved1  As Integer
    Reserved2  As Integer
    DIBSize    As Long
    DIBOffset  As Long
End Type

Private Type BITMAPFILEHEADER
    Type      As Integer
    Size      As Long
    Reserved1 As Integer
    Reserved2 As Integer
    OffBits   As Long
End Type

Private Type BITMAPINFOHEADER
    Size          As Long
    Width         As Long
    Height        As Long
    Planes        As Integer
    BitCount      As Integer
    Compression   As Long
    SizeImage     As Long
    XPelsPerMeter As Long
    YPelsPerMeter As Long
    ClrUsed       As Long
    ClrImportant  As Long
End Type

Private Type BITMAPINFO
    Header As BITMAPINFOHEADER
    Colors(1 To 256) As Long
End Type

Private Type TBITMAP
    Type As Long
    Width As Long
    Height As Long
    WidthBytes As Long
    Planes As Integer
    BitsPixel As Integer
    Bits As LongPtr
End Type

Private Type BITMAPDATA
    Width       As Long
    Height      As Long
    stride      As Long
    PixelFormat As Long
    scan0       As LongPtr
    Reserved    As LongPtr
End Type

Private Type RECT
    x      As Long
    y      As Long
    Width  As Long
    Height As Long
End Type

Private Type TRGBQuad
    Blue    As Byte
    Green   As Byte
    Red     As Byte
    Alpha   As Byte
End Type

Private Type TOLE_COLOR
    Long   As Long
End Type

Private Type TRGB
    Red     As Byte
    Green   As Byte
    Blue    As Byte
    None    As Byte
End Type


Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As LongPtr) As LongPtr
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal hDC As LongPtr) As Long
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hDC As LongPtr) As Long
Private Declare PtrSafe Function GetDIBits Lib "gdi32" (ByVal aHDC As LongPtr, ByVal hBitmapptr As LongPtr, ByVal nStartScan As Long, ByVal nNumScans As Long, ByRef lpBits As Any, ByRef lpBI As Any, ByVal wUsage As Long) As Long
Private Declare PtrSafe Function CreateDIBSection Lib "gdi32" (ByVal hDC As LongPtr, pbmi As Any, ByVal iUsage As Long, ByVal ppvBits As LongPtr, ByVal hSection As LongPtr, ByVal dwOffset As Long) As LongPtr
Private Declare PtrSafe Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As LongPtr
Private Declare PtrSafe Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As LongPtr, ByVal nWidth As Long, ByVal nHeight As Long) As LongPtr
Private Declare PtrSafe Function SelectObject Lib "gdi32.dll" (ByVal hDC As LongPtr, ByVal hgdiobj As LongPtr) As LongPtr
Private Declare PtrSafe Function DeleteObject Lib "gdi32.dll" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function StretchDIBits Lib "gdi32" (ByVal hDC As LongPtr, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, ByRef lpBits As LongPtr, ByRef lpBitsInfo As Any, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare PtrSafe Function StretchBlt Lib "gdi32" (ByVal hDC As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As LongPtr, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare PtrSafe Function SetStretchBltMode Lib "gdi32" (ByVal hDC As LongPtr, ByVal nStretchMode As Long) As Long
Private Declare PtrSafe Function SetBrushOrgEx Lib "gdi32" (ByVal hDC As LongPtr, ByVal nXOrg As Long, ByVal nYOrg As Long, lppt As Any) As Long

Private Declare PtrSafe Function BitBlt Lib "gdi32" (ByVal hDestDC As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As LongPtr, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare PtrSafe Function SetBkColor Lib "gdi32" (ByVal hDC As LongPtr, ByVal crColor As Long) As Long

Private Declare PtrSafe Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As LongPtr, ByVal nCount As Long, lpObject As Any) As Long
Private Declare PtrSafe Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As LongPtr, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare PtrSafe Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As LongPtr, ByVal dwCount As Long, lpBits As Any) As Long

Private Declare PtrSafe Function CLSIDFromString Lib "ole32" (ByVal lpszCLSID As LongPtr, ByRef pclsid As Guid) As Long
Private Declare PtrSafe Function IIDFromString Lib "ole32" (ByVal lpsz As LongPtr, ByRef lpiid As Any) As Long
Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32.dll" (ByRef PicDesc As Any, ByRef RefIID As Guid, ByVal fPictureOwnsHandle As Long, ByRef IPic As IPicture) As Long
Private Declare PtrSafe Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As LongPtr, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Private Declare PtrSafe Function GetHGlobalFromStream Lib "ole32" (ByVal Stream As IUnknown, ByRef hGlobal As LongPtr) As Long

Private Declare PtrSafe Function IStream_Size Lib "shlwapi" Alias "#214" (ByVal Stream As IUnknown, ByRef pui As Currency) As Long

Private Declare PtrSafe Function GdipSaveImageToStream Lib "gdiplus" (ByVal image As LongPtr, ByVal Stream As LongPtr, ByRef clsidEncoder As Guid, ByVal encoderParams As LongPtr) As Long
Private Declare PtrSafe Function GdipSaveImageToFile Lib "gdiplus.dll" (ByVal image As LongPtr, ByVal Filename As LongPtr, ByRef clsidEncoder As Guid, ByVal encoderParams As LongPtr) As Long
Private Declare PtrSafe Function GdipGetImageHeight Lib "gdiplus" (ByVal image As LongPtr, Height As Long) As Long
Private Declare PtrSafe Function GdipGetImageWidth Lib "gdiplus" (ByVal image As LongPtr, Width As Long) As Long
Private Declare PtrSafe Function GdipBitmapGetPixel Lib "gdiplus" (ByVal bitmap As LongPtr, ByVal x As Long, ByVal y As Long, ByRef Color As Long) As Long
Private Declare PtrSafe Function GdipBitmapSetPixel Lib "gdiplus" (ByVal bitmap As LongPtr, ByVal x As Long, ByVal y As Long, ByVal Color As Long) As Long
Private Declare PtrSafe Function GdipCreateBitmapFromScan0 Lib "gdiplus.dll" (ByVal nWidth As Long, ByVal Height As Long, ByVal stride As Long, ByVal PixelFormat As Long, ByRef scan0 As Any, ByRef nBitmap As LongPtr) As Long
Private Declare PtrSafe Function GdipCreateBitmapFromHBITMAP Lib "gdiplus.dll" (ByVal hbm As LongPtr, ByVal hPal As LongPtr, ByRef bitmap As LongPtr) As Long
Private Declare PtrSafe Function GdipCreateBitmapFromFile Lib "gdiplus" (ByVal Filename As LongPtr, bitmap As LongPtr) As Long
Private Declare PtrSafe Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" (ByVal bitmap As LongPtr, hbmReturn As LongPtr, ByVal background As Long) As Long
Private Declare PtrSafe Function GdipDisposeImage Lib "gdiplus" (ByVal image As LongPtr) As Long
Private Declare PtrSafe Function GdiplusStartup Lib "gdiplus" (ByRef token As LongPtr, ByRef inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As LongPtr = 0) As Long
Private Declare PtrSafe Function GdiplusShutdown Lib "gdiplus" (ByVal token As LongPtr) As Long
Private Declare PtrSafe Function GdipLoadImageFromFile Lib "gdiplus" (ByVal Filename As LongPtr, ByRef image As LongPtr) As Long
Private Declare PtrSafe Function GdipLoadImageFromStream Lib "gdiplus" (ByVal Stream As LongPtr, ByRef image As LongPtr) As Long
Private Declare PtrSafe Function GdipCreateSolidFill Lib "gdiplus" (ByVal pColor As Long, ByVal brush As LongPtr) As Long
Private Declare PtrSafe Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal image As LongPtr, graphics As LongPtr) As Long
Private Declare PtrSafe Function GdipFillRectangle Lib "gdiplus" (ByVal graphics As LongPtr, ByVal brush As LongPtr, ByVal x As Single, ByVal y As Single, ByVal nWidth As Single, ByVal nHeight As Single) As Long
Private Declare PtrSafe Function GdipSetSmoothingMode Lib "gdiplus" (ByVal mGraphics As LongPtr, ByVal mSmoothingMode As Long) As Long
Private Declare PtrSafe Function GdipDeleteBrush Lib "gdiplus" (ByVal mBrush As LongPtr) As Long
Private Declare PtrSafe Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As LongPtr) As Long
Private Declare PtrSafe Function GdipBitmapSetResolution Lib "gdiplus" (ByVal pbitmap As LongPtr, ByVal xdpi As Single, ByVal ydpi As Single) As Long
Private Declare PtrSafe Function GdipCreatePen1 Lib "gdiplus" (ByVal pColor As Long, ByVal Width As Long, ByVal unit As Long, ByRef hPen As LongPtr) As Long
Private Declare PtrSafe Function GdipDrawRectangle Lib "gdiplus" (ByVal hGraphics As LongPtr, ByVal hPen As LongPtr, ByVal x As Single, ByVal y As Single, ByVal nWidth As Single, ByVal nHeight As Single) As Long
Private Declare PtrSafe Function GdipDeletePen Lib "gdiplus" (ByVal hPen As LongPtr) As Long
Private Declare PtrSafe Function GdipCreateHICONFromBitmap Lib "gdiplus" (ByVal bitmap As LongPtr, ByRef hIcon As LongPtr) As Long
Private Declare PtrSafe Function GdipGetImageThumbnail Lib "gdiplus" (ByVal image As LongPtr, ByVal thumbWidth As Long, ByVal thumbHeight As Long, ByRef thumbImage As LongPtr, ByVal callback As Long, ByRef callbackData As Any) As Long
Private Declare PtrSafe Function GdipCreateBitmapFromGdiDib Lib "gdiplus" (ByRef gdiBitmapInfo As BITMAPINFOHEADER, ByRef gdiBitmapData As Any, bitmap As LongPtr) As Long
Private Declare PtrSafe Function GdipDrawImageRectI Lib "gdiplus.dll" (ByVal GpGraphics As LongPtr, ByVal GpImage As LongPtr, ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long) As Long
Private Declare PtrSafe Function GdipSetInterpolationMode Lib "gdiplus" (ByVal graphics As LongPtr, ByVal nInterpolationMode As Long) As Long
Private Declare PtrSafe Function GdipCreateBitmapFromGraphics Lib "gdiplus" (ByVal Width As Long, ByVal Height As Long, ByVal Target As LongPtr, bitmap As LongPtr) As Long
Private Declare PtrSafe Function GdipSetPixelOffsetMode Lib "gdiplus" (ByVal graphics As LongPtr, ByVal PixOffsetMode As Long) As Long
Private Declare PtrSafe Function GdipBitmapLockBits Lib "gdiplus" (ByVal bitmap As LongPtr, ByRef RECT As RECT, ByVal flags As Long, ByVal PixelFormat As Long, ByRef lockedBitmapData As Any) As Long
Private Declare PtrSafe Function GdipBitmapUnlockBits Lib "gdiplus" (ByVal bitmap As LongPtr, ByRef lockedBitmapData As Any) As Long
Private Declare PtrSafe Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As LongPtr, ByRef graphics As LongPtr) As Long

Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
Private Declare PtrSafe Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As LongPtr
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As LongPtr
Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)

Private Declare PtrSafe Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As LongPtr, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As LongPtr) As LongPtr
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function ReadFile Lib "kernel32" (ByVal hFile As LongPtr, ByRef Buffer As Any, ByVal nNumberOfBytesToWrite As Long, ByRef NumberOfBytesWritten As Long, ByVal lpOverlapped As LongPtr) As Long
Private Declare PtrSafe Function WriteFile Lib "kernel32" (ByVal hFile As LongPtr, ByRef Buffer As Any, ByVal nNumberOfBytesToWrite As Long, ByRef NumberOfBytesWritten As Long, ByVal lpOverlapped As LongPtr) As Long
Private Declare PtrSafe Function GetFileSize Lib "kernel32" (ByVal hFile As LongPtr, ByRef FileSizeHigh As Long) As Long

Private Enum ImageLockMode
    ImageLockModeRead = 1
    ImageLockModeWrite = 2
    ImageLockModeUserInputBuf = 4
End Enum

Private Enum InterpolationMode
    InterpolationModeInvalid = -1
    InterpolationModeDefault = 0&
    InterpolationModeLowQuality = 1&
    InterpolationModeHighQuality = 2&
    InterpolationModeBilinear = 3&
    InterpolationModeBicubic = 4&
    InterpolationModeNearestNeighbor = 5&
    InterpolationModeHighQualityBilinear = 6&
    InterpolationModeHighQualityBicubic = 7&
End Enum

Private Const PixelOffsetModeHalf = 4
Private Const PixelFormat32bppARGB = &H26200A
Private Const IID_IPictureDisp As String = "{7BF80981-BF32-101A-8BBB-00AA00300CAB}"
Private Const CLSID_Suffix As String = "-1A04-11D3-9A73-0000F81EF32E}"
Private Const CLSID_BMP    As String = "{557CF400" & CLSID_Suffix
Private Const CLSID_JPG    As String = "{557CF401" & CLSID_Suffix
Private Const CLSID_GIF    As String = "{557CF402" & CLSID_Suffix
Private Const CLSID_EMF    As String = "{557CF403" & CLSID_Suffix
Private Const CLSID_WMF    As String = "{557CF404" & CLSID_Suffix
Private Const CLSID_TIF    As String = "{557CF405" & CLSID_Suffix
Private Const CLSID_PNG    As String = "{557CF406" & CLSID_Suffix
Private Const CLSID_ICO    As String = "{557CF407" & CLSID_Suffix

Private Const S_OK = 0
Private Const PICTYPE_BITMAP = 1
Private Const PICTYPE_METAFILE = 2
Private Const PICTYPE_ICON = 3
Private Const PICTYPE_ENHMETAFILE = 4

Private Const GMEM_MOVEABLE As Long = &H2&

Private Const DIB_RGB_COLORS = 0&
Private Const SRCCOPY = &HCC0020
Private Const HALFTONE = 4
Private Const UNITPIXEL = 2

'ファイル関係
Private Const GENERIC_WRITE = &H40000000
Private Const GENERIC_READ = &H80000000
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const CREATE_ALWAYS = 2
Private Const OPEN_EXISTING = 3
Private Const OPEN_ALWAYS = 4
Private Const INVALID_HANDLE_VALUE = -1

'クリップボード関係
Private Const CF_BITMAP = 2
Private Const CF_DIB = 8
Private Const CF_DIF = 5
Private Const CF_DSPBITMAP = &H82
Private Const CF_DSPENHMETAFILE = &H8E
Private Const CF_DSPMETAFILEPICT = &H83
Private Const CF_DSPTEXT = &H81
Private Const CF_ENHMETAFILE = 14
Private Const CF_METAFILEPICT = 3
Private Const CF_PALETTE = 9
Private Const CF_PENDATA = 10
Private Const CF_TEXT = 1
Private Const CF_TIFF = 6
Private Const CF_UNICODETEXT = 13

'*****************************************************************************
'クラス内ユーザー定義
'*****************************************************************************
Private FPixels() As TRGBQuad
Private FWidth As Long
Private FHeight As Long
Private FGdiPlus As LongPtr

Private Const MAX_WIDTH = 64
Private Const MAX_HEIGHT = 64

'*****************************************************************************
'[概要] コンストラクタ
'*****************************************************************************
Private Sub Class_Initialize()
    Dim GdiInput As GdiplusStartupInput
    GdiInput.GdiplusVersion = 1&
    If GdiplusStartup(FGdiPlus, GdiInput) <> S_OK Then
        Call Err.Raise(513, "GdiplusStartupエラー")
    End If
End Sub

'*****************************************************************************
'[概要] デストラクタ
'*****************************************************************************
Private Sub Class_Terminate()
    If FGdiPlus <> 0 Then Call GdiplusShutdown(FGdiPlus)
End Sub

'*****************************************************************************
'[概要] Cellの色を取得する
'[引数] 対象のセル
'[戻値] TRGBQuad
'*****************************************************************************
Private Function CellToColor(ByRef objCell As Range) As TRGBQuad
    Dim strNumeric As String
    Dim Alpha As Byte
    With objCell.Interior
        If .ColorIndex = xlColorIndexAutomatic Or .ColorIndex = xlColorIndexNone Then
            '透明
            CellToColor = OleColorToARGB(&HFFFFFF, 0)
        Else
            Alpha = &HFF '不透明
            '半透明
            If .Pattern = xlGray8 Then
                strNumeric = Replace(objCell.Value, "$", "&H", 1, 1)
                If IsNumeric(strNumeric) Then
                    If 0 <= CInt(strNumeric) And CInt(strNumeric) <= 256 Then
                        'セルに入力された数値がアルファ値
                        Alpha = CInt(strNumeric)
                    End If
                End If
            End If
            CellToColor = OleColorToARGB(.Color, Alpha)
        End If
    End With
End Function

'*****************************************************************************
'[概要] Cellの色を取得する
'[引数] 対象のセル
'[戻値] TRGBQuad
'*****************************************************************************
Public Function CellToColor2(ByRef objCell As Range) As Long
    CellToColor2 = ARGBToLong(CellToColor(objCell))
End Function

'*****************************************************************************
'[概要] Cellの色を設定する
'[引数] 対象のセル，設定する色
'[戻値] TRGBQuad
'*****************************************************************************
Private Sub ColorToCell(ByRef objCell As Range, ByRef Color As TRGBQuad)
    With objCell.Interior
        Select Case Color.Alpha
        Case 0   '透明
            .Pattern = xlGray8
            .PatternColorIndex = xlAutomatic
            .TintAndShade = 0
            .PatternTintAndShade = 0
            .ColorIndex = xlAutomatic
            objCell.Value = ""
        Case 255 '不透明
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .TintAndShade = 0
            .PatternTintAndShade = 0
            .Color = ARGBToRGB(Color)
            objCell.Value = ""
        Case Else '半透明
            .Pattern = xlGray8
            .PatternColor = &HFFFFFF
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .PatternTintAndShade = 0
            .Color = ARGBToRGB(Color)
            With objCell
                .Value = Color.Alpha
                .Font.Color = ARGBToRGB(Color) '文字を背景色と同じにする
            End With
        End Select
    End With
End Sub

'*****************************************************************************
'[概要] TRGBQuad型をOLE_COLORに変換する
'[引数] TRGBQuad
'[戻値] OLE_COLOR
'*****************************************************************************
Private Function ARGBToRGB(ByRef ARGB As TRGBQuad) As OLE_COLOR
    Dim Color As TOLE_COLOR
    Dim RGB As TRGB
    With RGB
        .Red = ARGB.Red
        .Green = ARGB.Green
        .Blue = ARGB.Blue
    End With
    LSet Color = RGB
    ARGBToRGB = Color.Long
End Function

'*****************************************************************************
'[概要] TRGBQuad型をLongに変換する
'[引数] TRGBQuad
'[戻値] OLE_COLOR
'*****************************************************************************
Private Function ARGBToLong(ByRef ARGB As TRGBQuad) As Long
    Dim Color As TOLE_COLOR
    LSet Color = ARGB
    ARGBToLong = Color.Long
End Function

'*****************************************************************************
'[概要] TRGBQuad型をLongに変換する
'[引数] TRGBQuad
'[戻値] OLE_COLOR
'*****************************************************************************
Private Function LongToARGB(ByVal Value As Long) As TRGBQuad
    Dim Color As TOLE_COLOR
    Color.Long = Value
    LSet LongToARGB = Color
End Function

'*****************************************************************************
'[概要] OLE_COLORをTRGBQuad型に変換する
'[引数] OLE_COLOR，アルファ値(省略時は透過なし)
'[戻値] TRGBQuad
'*****************************************************************************
Private Function OleColorToARGB(ByVal lngColor As OLE_COLOR, Optional Alpha As Byte = 255) As TRGBQuad
    Dim RGB As TRGB
    Dim Color As TOLE_COLOR
    Color.Long = lngColor
    LSet RGB = Color
    With RGB
        OleColorToARGB = RGBToARGB(.Red, .Green, .Blue, Alpha)
    End With
End Function

'*****************************************************************************
'[概要] RGB & アルファ値をTRGBQuad型に変換する
'[引数] RGB，アルファ値(省略時は透過なし)
'[戻値] TRGBQuad
'*****************************************************************************
Private Function RGBToARGB(ByVal Red As Byte, ByVal Green As Byte, ByVal Blue As Byte, Optional Alpha As Byte = 255) As TRGBQuad
    With RGBToARGB
        .Red = Red
        .Green = Green
        .Blue = Blue
        .Alpha = Alpha
    End With
End Function

'*****************************************************************************
'[概要] 指定色(-1は透明とする)とRGBQuad色が同一かどうか判定する
'[引数] 指定色(-1は透明とする)，RGBQuad
'[戻値] True:同一
'*****************************************************************************
Private Function SameColor(ByVal lngColor As Long, RGBQuad As TRGBQuad) As Boolean
    If lngColor = -1 Then
        SameColor = (RGBQuad.Alpha = 0)
    ElseIf RGBQuad.Alpha = 0 Then
        SameColor = False
    Else
        With OleColorToARGB(lngColor)
            SameColor = (.Red = RGBQuad.Red) And _
                        (.Green = RGBQuad.Green) And _
                        (.Blue = RGBQuad.Blue)
        End With
    End If
End Function

'*****************************************************************************
'[概要] 画像の幅と高さを設定
'[引数] なし
'*****************************************************************************
Public Function SetSize(ByVal lngWidth As Long, ByVal lngHeight As Long)
    ReDim FPixels(1 To lngWidth, 1 To lngHeight)
    FWidth = lngWidth
    FHeight = lngHeight
End Function

'*****************************************************************************
'[概要] 画像の幅
'[引数] なし
'*****************************************************************************
Public Property Get Width() As Long
    Width = FWidth
End Property

'*****************************************************************************
'[概要] 画像の高さ
'[引数] なし
'*****************************************************************************
Public Property Get Height() As Long
    Height = FHeight
End Property

'*****************************************************************************
'[概要] 行と列を置換する
'[引数] なし
'[戻値] なし
'*****************************************************************************
Public Sub Transpose()
    ReDim Work(1 To FHeight, 1 To FWidth) As TRGBQuad
    Dim x As Long, y As Long
    For x = 1 To FHeight
        For y = 1 To FWidth
            Work(x, y) = FPixels(y, x)
        Next
    Next
    Dim Swap As Long
    Swap = FWidth
    FWidth = FHeight
    FHeight = Swap
    FPixels = Work
End Sub

'*****************************************************************************
'[概要] 水平方向に画像を反転させる
'[引数] なし
'[戻値] なし
'*****************************************************************************
Public Sub FlipHorizontal()
    ReDim Work(1 To FWidth, 1 To FHeight) As TRGBQuad
    Dim x As Long, y As Long
    For x = 1 To FWidth
        For y = 1 To FHeight
            Work(x, y) = FPixels(x, FHeight - y + 1)
        Next
    Next
    FPixels = Work
End Sub

'*****************************************************************************
'[概要] 垂直方向に画像を反転させる
'[引数] なし
'[戻値] なし
'*****************************************************************************
Public Sub FlipVertical()
    Call Transpose
    Call FlipHorizontal
    Call Transpose
End Sub

'*****************************************************************************
'[概要] 画像を回転させる
'[引数] 角度:90,180,270(-90)
'[戻値] なし
'*****************************************************************************
Public Sub Rotate(ByVal Angle As Long)
    Select Case Angle
    Case 90
        Call FlipHorizontal
        Call Transpose
    Case 180
        Call FlipVertical
        Call FlipHorizontal
    Case 270, -90
        Call Transpose
        Call FlipHorizontal
    End Select
End Sub

'*****************************************************************************
'[概要] GUID文字列からGUIDを生成
'[引数] GUID文字列
'[戻値] GUID
'*****************************************************************************
Private Function GetCLSID(ByVal strGuid As String) As Guid
    Call CLSIDFromString(StrPtr(strGuid), GetCLSID)
End Function

'*****************************************************************************
'[概要] ファイルから画像を読込む
'[引数] ファイル名
'[戻値] なし
'*****************************************************************************
Private Sub LoadImageFromFile(ByVal strFile As String)
On Error GoTo Finalization
    Dim hGdiImage As LongPtr
    If GdipLoadImageFromFile(ByVal StrPtr(strFile), hGdiImage) <> S_OK Then
        Call Err.Raise(513, "GdipLoadImageFromFileエラー")
    End If
    Call GdipBitmapSetResolution(hGdiImage, 96#, 96#) '標準のDPIに設定
    Call GetPixelsFromBitmap(hGdiImage)
    
    '幅・高さの最大値を設定する（パフォーマンスを確保するため）
    If FWidth > MAX_WIDTH Then FWidth = MAX_WIDTH
    If FHeight > MAX_HEIGHT Then FHeight = MAX_HEIGHT
Finalization:
    If hGdiImage <> 0 Then Call GdipDisposeImage(hGdiImage)
    If Err.Number <> 0 Then Call Err.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'*****************************************************************************
'[概要] ファイルに画像を保存する
'[引数] ファイル名
'[戻値] なし
'*****************************************************************************
Public Function SaveImageToFile(ByVal strFile As String) As Boolean
On Error GoTo Finalization
    Dim strGuid As String

    Select Case GetFileExtension(strFile)
    Case "ICO"
        SaveImageToFile = SaveToIcon(strFile) 'アルファチェネルの設定あり
        Exit Function
    Case "BMP"
        SaveImageToFile = SaveToBmp(strFile) 'アルファチェネルの設定あり
        Exit Function
    Case "PNG"
        strGuid = CLSID_PNG
    Case Else
        strGuid = CLSID_BMP
    End Select
    
    Dim hGdiImage As LongPtr
    If GdipCreateBitmapFromScan0(FWidth, FHeight, 0, PixelFormat32bppARGB, ByVal 0&, hGdiImage) <> S_OK Then
        Call Err.Raise(513, "GdipCreateBitmapFromScan0エラー")
    End If
    Call SetPixelsToBitmap(hGdiImage)
    Call GdipBitmapSetResolution(hGdiImage, 96#, 96#) '標準のDPIに設定
    If GdipSaveImageToFile(hGdiImage, ByVal StrPtr(strFile), GetCLSID(strGuid), 0) <> S_OK Then
        Call Err.Raise(513, "GdipSaveImageToFileエラー")
    End If
    Call GdipDisposeImage(hGdiImage)
    SaveImageToFile = True
    Exit Function
Finalization:
    If hGdiImage <> 0 Then Call GdipDisposeImage(hGdiImage)
    If Err.Number <> 0 Then Call Err.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Function

'*****************************************************************************
'[概要] Bitmapハンドルから画像を取込む
'       ※アルファチャネルが反映されないため却下
'[引数] Bitmapハンドル
'[戻値] なし
'*****************************************************************************
'Public Sub GetPixelsFromHBITMAP(ByVal hBitmap As LongPtr)
'On Error GoTo Finalization
'    Dim hGdiImage As LongPtr
'    If GdipCreateBitmapFromHBITMAP(hBitmap, 0&, hGdiImage) <> S_OK Then
'        Call Err.Raise(513, "GdipCreateBitmapFromHBITMAPエラー")
'    End If
'    Call GetPixelsFromBitmap(hGdiImage)
'Finalization:
'    If hGdiImage <> 0 Then Call GdipDisposeImage(hGdiImage)
'    If Err.Number <> 0 Then Call Err.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
'End Sub

'*****************************************************************************
'[概要] Bitmapオブジェクトから画像を取込む
'[引数] Bitmapオブジェクト,縮小する幅と高さ
'[戻値] なし
'*****************************************************************************
Private Sub GetPixelsFromBitmap(ByVal hGdiImage As LongPtr)
    Call GdipGetImageWidth(hGdiImage, FWidth)
    Call GdipGetImageHeight(hGdiImage, FHeight)
    ReDim FPixels(1 To FWidth, 1 To FHeight)
    Call GetPixelsFromImage(hGdiImage)
End Sub

'*****************************************************************************
'[概要] 画像のサイズを変更する
'[引数] 新しいサイズ
'[戻値] なし
'*****************************************************************************
Public Sub Resize(ByVal lngWidth As Long, ByVal lngHeight As Long)
    Dim hNewGraphics As LongPtr
    Dim hOldBmp      As LongPtr
    Dim hNewBmp      As LongPtr
    
    If FWidth = lngWidth And FHeight = lngHeight Then
        '幅・高さの最大値を設定する（パフォーマンスを確保するため）
        If FWidth > MAX_WIDTH Then FWidth = MAX_WIDTH
        If FHeight > MAX_HEIGHT Then FHeight = MAX_HEIGHT
        Exit Sub
    End If

'    Dim hSrcDC As LongPtr
'    Dim hDstDc As LongPtr
'    Dim hSrcBitmap As LongPtr
'    Dim hDstBitmap As LongPtr
'    hSrcDC = CreateCompatibleDC(0)
'    If GdipCreateBitmapFromScan0(FWidth, FHeight, 0, PixelFormat32bppARGB, ByVal 0&, hOldBmp) <> S_OK Then
'        Call Err.Raise(513, "GdipCreateBitmapFromScan0エラー")
'    End If
'    Call SetPixelsToBitmap(hOldBmp)
'    Call GdipCreateHBITMAPFromBitmap(hOldBmp, hSrcBitmap, 0)
'    hSrcDC = CreateCompatibleDC(0)
'    Call SelectObject(hSrcDC, hSrcBitmap)
'
'    hDstBitmap = CreateBitmap(lngWidth, lngHeight, 1, 32, ByVal 0&)
'    hDstDc = CreateCompatibleDC(0)
'    Call SelectObject(hDstDc, hDstBitmap)
''    Call SetBrushOrgEx(hDstDc, 0, 0, ByVal 0)
''    Call SetStretchBltMode(hDstDc, HALFTONE)
'    Call StretchBlt(hDstDc, 0, 0, lngWidth, lngHeight, hSrcDC, 0, 0, FWidth, FHeight, SRCCOPY)
'    Call SetSize(lngWidth, lngHeight)
'    Call GetPixelsFromHBITMAP(hDstBitmap)
'    Exit Sub
    
'    Dim hDC As LongPtr
'    hDC = CreateCompatibleDC(0)
'    Call GdipCreateFromHDC(hDC, hGraphics)
'    Call DeleteDC(hDC)
    
On Error GoTo Finalization
    If GdipCreateBitmapFromScan0(FWidth, FHeight, 0, PixelFormat32bppARGB, ByVal 0&, hOldBmp) <> S_OK Then
        Call Err.Raise(513, "GdipCreateBitmapFromScan0エラー")
    End If
    Call SetPixelsToBitmap(hOldBmp)
    
    If GdipCreateBitmapFromScan0(lngWidth, lngHeight, 0, PixelFormat32bppARGB, ByVal 0&, hNewBmp) <> S_OK Then
        Call Err.Raise(513, "GdipCreateBitmapFromScan0エラー")
    End If
    If GdipGetImageGraphicsContext(hNewBmp, hNewGraphics) = 0 Then
'        Call GdipSetInterpolationMode(hGraphics, InterpolationModeHighQuality)
        Call GdipSetInterpolationMode(hNewGraphics, InterpolationModeNearestNeighbor)
        Call GdipSetPixelOffsetMode(hNewGraphics, PixelOffsetModeHalf)
        Call GdipDrawImageRectI(hNewGraphics, hOldBmp, 0, 0, lngWidth, lngHeight)
    End If
    
    Call SetSize(lngWidth, lngHeight)
    Call GetPixelsFromImage(hNewBmp)
    
    '幅・高さの最大値を設定する（パフォーマンスを確保するため）
    If FWidth > MAX_WIDTH Then FWidth = MAX_WIDTH
    If FHeight > MAX_HEIGHT Then FHeight = MAX_HEIGHT
Finalization:
    If hNewGraphics <> 0 Then Call GdipDeleteGraphics(hNewGraphics)
    If hNewBmp <> 0 Then Call GdipDisposeImage(hNewBmp)
    If hOldBmp <> 0 Then Call GdipDisposeImage(hOldBmp)
End Sub

'*****************************************************************************
'[概要] BitmapオブジェクトからFPixels()を設定
'[引数] Bitmapオブジェクト
'[戻値] なし
'*****************************************************************************
Private Sub GetPixelsFromImage(ByVal hGdiImage As LongPtr)
    Dim uRect As RECT
    Dim BmpData As BITMAPDATA
    
    With uRect
        .x = 0
        .y = 0
        .Width = FWidth
        .Height = FHeight
    End With
    
    With BmpData
        .Width = FWidth
        .Height = FHeight
        .PixelFormat = PixelFormat32bppARGB
        .scan0 = VarPtr(FPixels(1, 1))
        .stride = FWidth * 4
    End With
    
    If GdipBitmapLockBits(hGdiImage, uRect, ImageLockModeUserInputBuf Or ImageLockModeRead, PixelFormat32bppARGB, BmpData) <> S_OK Then
        Call Err.Raise(513, "GdipBitmapLockBitsエラー")
    End If
    Call GdipBitmapUnlockBits(hGdiImage, BmpData)
End Sub

'*****************************************************************************
'[概要] Bitmapオブジェクトに画像を設定 ※アルファチャネルあり
'[引数] Bitmapオブジェクト
'[戻値] なし
'*****************************************************************************
Private Sub SetPixelsToBitmap(ByRef hGdiImage As LongPtr)
    Dim x As Long, y As Long
    Dim Color As TOLE_COLOR
    For x = 1 To FWidth
        For y = 1 To FHeight
            LSet Color = FPixels(x, y)
            Call GdipBitmapSetPixel(hGdiImage, x - 1, y - 1, Color.Long)
        Next
    Next
End Sub

'*****************************************************************************
'[概要] 画像を読込む
'[引数] ファイル名
'[戻値] True:成功
'*****************************************************************************
Public Function LoadFromFile(ByVal strFilename As String) As Boolean
On Error GoTo ErrHandle
    Call LoadImageFromFile(strFilename)
    LoadFromFile = True
ErrHandle:
End Function

'*****************************************************************************
'[概要] 画像を保存する
'[引数] ファイル名
'[戻値] True:成功
'*****************************************************************************
Public Function SaveToFile(ByVal strFilename As String) As Boolean
On Error GoTo ErrHandle
    Call SaveImageToFile(strFilename)
    SaveToFile = True
ErrHandle:
End Function

''*****************************************************************************
''[概要] PNG形式の画像をクリップボードから読込む
''[引数] 縮小する幅と高さ
''[戻値] True:成功
''*****************************************************************************
'Public Function LoadPNGFromClipbord() As Boolean
'    Dim CF_PNG  As Long
'    Dim hGlobal As LongPtr
'
'    If OpenClipboard(0) = 0 Then Exit Function
'On Error GoTo Finalization
'    CF_PNG = RegisterClipboardFormat("PNG")
'    hGlobal = GetClipboardData(CF_PNG)
'    If hGlobal = 0 Then GoTo Finalization
'
'    Dim Stream    As IUnknown
'    Dim hGdiImage As LongPtr
'    If CreateStreamOnHGlobal(ByVal hGlobal, 0, Stream) <> S_OK Then GoTo Finalization
'    If GdipLoadImageFromStream(ObjPtr(Stream), hGdiImage) <> S_OK Then
'        Call Err.Raise(513, "GdipLoadImageFromStreamエラー")
'    End If
'    Call GdipBitmapSetResolution(hGdiImage, 96#, 96#) '標準のDPIに設定
'    Call GetPixelsFromBitmap(hGdiImage)
'    LoadPNGFromClipbord = True
'Finalization:
'    Call CloseClipboard
'    If hGdiImage <> 0 Then Call GdipDisposeImage(hGdiImage)
'    If Err.Number <> 0 Then Call Err.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
'End Function

'*****************************************************************************
'[概要] PNG形式の画像をクリップボードに書込む
'[引数] なし
'[戻値] True:成功
'*****************************************************************************
Private Function SavePNGToClipbord() As Boolean
On Error GoTo ErrHandle
    Dim Stream    As IUnknown
    Dim hGdiImage As LongPtr
    If CreateStreamOnHGlobal(0, 1, Stream) <> S_OK Then GoTo ErrHandle
    '32Bitカラーのbitmapオブジェクトを生成
    If GdipCreateBitmapFromScan0(FWidth, FHeight, 0, PixelFormat32bppARGB, ByVal 0&, hGdiImage) <> S_OK Then
        Call Err.Raise(513, "GdipCreateBitmapFromScan0エラー")
    End If
    Call SetPixelsToBitmap(hGdiImage)
    Call GdipBitmapSetResolution(hGdiImage, 96#, 96#) '標準のDPIに設定
    If GdipSaveImageToStream(hGdiImage, ByVal ObjPtr(Stream), GetCLSID(CLSID_PNG), 0) <> S_OK Then
        Call Err.Raise(513, "GdipSaveImageToStreamエラー")
    End If
    
    Dim hGlobal As LongPtr
    If GetHGlobalFromStream(Stream, hGlobal) <> S_OK Then GoTo ErrHandle
    If hGlobal = 0 Then GoTo ErrHandle
    
    Dim CF_PNG As Long
    CF_PNG = RegisterClipboardFormat("PNG")
    If SetClipboardData(CF_PNG, hGlobal) = 0 Then GoTo ErrHandle
    
    SavePNGToClipbord = True
ErrHandle:
    If hGdiImage <> 0 Then Call GdipDisposeImage(hGdiImage)
End Function

'*****************************************************************************
'[概要] BMP形式の画像をクリップボードから読込む
'[引数] なし
'[戻値] True:成功
'*****************************************************************************
Public Function LoadBMPFromClipbord() As Boolean
    Dim hBitmap As LongPtr
    Dim hGdiImage As LongPtr
    
    If OpenClipboard(0) = 0 Then Exit Function
    hBitmap = GetClipboardData(CF_BITMAP)
    If hBitmap = 0 Then
        Call CloseClipboard
        Exit Function
    End If
    Call GetPixelsFromHBITMAP(hBitmap)
        
    LoadBMPFromClipbord = True
Finalization:
    If hGdiImage <> 0 Then Call GdipDisposeImage(hGdiImage)
    Call CloseClipboard
End Function

'*****************************************************************************
'[概要] FPixelsからアルファチャネル有のBitmapオブジェクトを作成する
'       アルファチャネルが反映されないため却下
'[引数] なし
'[戻値] Bitmapオブジェクト
'*****************************************************************************
'Private Function CreateBitmapFromPixels() As LongPtr
'On Error GoTo Finalization
'    Dim BmpInfHeader  As BITMAPINFOHEADER
'
'    With BmpInfHeader
'        .Size = Len(BmpInfHeader)
'        .Width = FWidth
'        .Height = -FHeight 'ボトムアップからトップダウンに変更
'        .Planes = 1 '常に1
'        .BitCount = 32 'Bit
'        .Compression = 0 'BI_RGB(無圧縮)
'        .SizeImage = 0 '省略可
'    End With
'
'    Dim hGdiImage As LongPtr
'    'アルファチャネルは反映されない
'    Call GdipCreateBitmapFromGdiDib(BmpInfHeader, FPixels(1, 1), hGdiImage)
'    CreateBitmapFromPixels = hGdiImage
'Finalization:
'End Function

'*****************************************************************************
'[概要] BMP形式の画像をクリップボードに書込む
'[引数] なし
'[戻値] True:成功
'*****************************************************************************
Private Function SaveBMPToClipbord() As Boolean
    Dim hGdiImage As LongPtr
    Dim hBitmap As LongPtr
On Error GoTo Finalization
    If GdipCreateBitmapFromScan0(FWidth, FHeight, 0, PixelFormat32bppARGB, ByVal 0&, hGdiImage) <> S_OK Then
        Exit Function
    End If
    Call SetPixelsToBitmap(hGdiImage)
    If GdipCreateHBITMAPFromBitmap(hGdiImage, hBitmap, 0) <> S_OK Then GoTo Finalization
    If SetClipboardData(CF_BITMAP, hBitmap) = 0 Then GoTo Finalization
    
    SaveBMPToClipbord = True
Finalization:
    If hGdiImage <> 0 Then Call GdipDisposeImage(hGdiImage)
End Function

'*****************************************************************************
'[概要] PNG形式とBMP形式の画像をクリップボードに書込む
'[引数] なし
'[戻値] True:成功
'*****************************************************************************
Public Function SaveImageToClipbord() As Boolean
    If OpenClipboard(0) = 0 Then Exit Function
On Error GoTo Finalization
    Call EmptyClipboard
    SaveImageToClipbord = SavePNGToClipbord() Or SaveBMPToClipbord()
Finalization:
    Call CloseClipboard
End Function

'*****************************************************************************
'[概要] 画素の色を置換する
'[引数] 置換前の色のセル，置換後の色のセル
'[戻値] なし
'*****************************************************************************
Public Sub ChangeColor(ByVal objSrcCell As Range, ByVal objDstCell As Range)
    Dim x As Long, y As Long
    For x = 1 To FWidth
        For y = 1 To FHeight
            If SameColor2(CellToColor(objSrcCell), FPixels(x, y)) Then
                FPixels(x, y) = CellToColor(objDstCell)
            End If
        Next
    Next
End Sub

'*****************************************************************************
'[概要] 画素の色を置換する
'[引数] 置換前の色のセル，置換対象のセル
'[戻値] なし
'*****************************************************************************
Public Sub ChangeCellColor(ByVal objSrcCell As Range, ByRef objDstCell As Range)
    Call ColorToCell(objDstCell, CellToColor(objSrcCell))
End Sub

'*****************************************************************************
'[概要] Cellの色からFPixelsを設定
'[引数] Cellの範囲
'[戻値] なし
'*****************************************************************************
Public Sub GetPixelsFromRange(ByRef objRange As Range)
    Call SetSize(objRange.Columns.Count, objRange.Rows.Count)
    Dim x As Long, y As Long
    For x = 1 To FWidth
        For y = 1 To FHeight
            FPixels(x, y) = CellToColor(objRange.Cells(y, x))
        Next
    Next
End Sub

'*****************************************************************************
'[概要] FPixelsからCellに色を設定
'[引数] 色を設定するセル，Mode 1:特定色を対象,2:特定色以外を対象
'       特定色指定時のセル，対象外のセルを透明とするかどうか
'[戻値] なし
'*****************************************************************************
Public Sub SetPixelsToRange(ByVal objRange As Range, Optional Mode As Long = 0, Optional objColorCell As Range = Nothing, Optional blnTransparent As Boolean = False)
    Dim objRange2 As Range
    Set objRange2 = objRange.Resize(FHeight, FWidth)
    Dim objCell As Range
    
    Application.ScreenUpdating = False
    Dim x As Long, y As Long
    For x = 1 To FWidth
        For y = 1 To FHeight
            Set objCell = objRange.Cells(y, x)
            Select Case Mode
            Case 0
                Call ColorToCell(objCell, FPixels(x, y))
            Case 1, 2
                If SameColor2(CellToColor(objColorCell), FPixels(x, y)) = (Mode = 1) Then
                    Call ColorToCell(objCell, FPixels(x, y))
                Else
                    If blnTransparent Then
                        Call ColorToCell(objCell, OleColorToARGB(&HFFFFFF, 0))
                    End If
                End If
            End Select
        Next
    Next
End Sub

'*****************************************************************************
'[概要] アイコンファイルとして画像を保存する
'[引数] ファイル名
'[戻値] True:成功
'*****************************************************************************
Public Function SaveToIcon(ByVal strFilename As String) As Boolean
On Error GoTo Finalization
    Dim MaskSize    As Long
    Dim MaskBits()  As Byte
    
    '左下から上に向かって設定するため水平方向に反転させる
    Call FlipHorizontal
    MaskSize = MakeMaskBits(MaskBits())
    
    Dim IconFileHeader As ICONDIR
    Dim IconInfHeader  As ICONDIRENTRY
    Dim BmpInfHeader   As BITMAPINFOHEADER
    
    With IconFileHeader
        .Type = 1 'アイコン
        .Count = 1
    End With
    
    With IconInfHeader
        .Width = FWidth
        .Height = FHeight
        .ColorCount = 0 '256色以上なら0
        .DIBSize = Len(BmpInfHeader) + FWidth * 4 * FHeight + MaskSize '32Bitの時
        .DIBOffset = Len(IconFileHeader) + Len(IconInfHeader)
    End With
    
    With BmpInfHeader
        .Size = Len(BmpInfHeader)
        .Width = FWidth
        .Height = FHeight * 2
        .Planes = 1 '常に1
        .BitCount = 32 'Bit
        .Compression = 0 'BI_RGB(無圧縮)
        .SizeImage = 0 '省略可
    End With
    
'    Dim x, y
'    For x = 1 To FWidth
'        For y = 1 To FHeight
'            FPixels2(x, y).Alpha = 0 '検証用
'        Next
'    Next
    
    Dim File As Integer
    Dim blnOpen As Boolean
    File = FreeFile()
    Open strFilename For Binary Access Write As #File
    blnOpen = True
    Put #File, , IconFileHeader
    Put #File, , IconInfHeader
    Put #File, , BmpInfHeader
    Put #File, , FPixels()
    Put #File, , MaskBits
    
    SaveToIcon = True
Finalization:
    If blnOpen Then Close #File
    Call FlipHorizontal  '水平方向の反転を戻す
End Function

'*****************************************************************************
'[概要] なぜか減色されてしまうため使えず
'*****************************************************************************
Private Function SaveToIcon2(ByVal strFilename As String) As Boolean
    Dim uPicInfo As PICTDESC_ICON
    Dim hGdiImage As LongPtr
    Dim hIcon As LongPtr
    Dim Icon As IPicture

    If GdipCreateBitmapFromScan0(FWidth, FHeight, 0, PixelFormat32bppARGB, ByVal 0&, hGdiImage) <> S_OK Then Exit Function
On Error GoTo Finalization
    Call SetPixelsToBitmap(hGdiImage)
    Call GdipCreateHICONFromBitmap(hGdiImage, hIcon)
    With uPicInfo
        .Size = Len(uPicInfo)
        .Type = PICTYPE_ICON
        .hIcon = hIcon
    End With

    Call GdipCreateHICONFromBitmap(hGdiImage, hIcon)
    Call OleCreatePictureIndirect(uPicInfo, GetCLSID(IID_IPictureDisp), True, Icon)
    Call SavePicture(Icon, strFilename)

    SaveToIcon2 = True
Finalization:
    If hGdiImage <> 0 Then Call GdipDisposeImage(hGdiImage)
End Function

'*****************************************************************************
'[概要] FPixelsからMaskBitsを作成
'[引数] 設定するMaskBits
'[戻値] 必要バイト数
'*****************************************************************************
Private Function MakeMaskBits(ByRef MaskBits() As Byte) As Long
    Dim WidhBit As Long
    WidhBit = Int((FWidth + 31) / 32) * 32 '1行ごとの幅は4バイトの倍数にする
    Dim x  As Long
    Dim y  As Long
    ReDim MaskWork(1 To WidhBit * FHeight) As Byte 'Bit
    
    For y = 0 To FHeight - 1
        For x = 1 To FWidth
            If FPixels(x, y + 1).Alpha = 0 Then
                MaskWork(y * WidhBit + x) = 1
            End If
        Next
    Next
    
    MakeMaskBits = WidhBit / 8 * FHeight '必要バイト数
    ReDim MaskBits(0 To MakeMaskBits - 1)
    Dim i As Long
    For i = 0 To MakeMaskBits - 1
        MaskBits(i) = MaskWork(i * 8 + 8) _
                    + MaskWork(i * 8 + 7) * 2 _
                    + MaskWork(i * 8 + 6) * 4 _
                    + MaskWork(i * 8 + 5) * 8 _
                    + MaskWork(i * 8 + 4) * 16 _
                    + MaskWork(i * 8 + 3) * 32 _
                    + MaskWork(i * 8 + 2) * 64 _
                    + MaskWork(i * 8 + 1) * 128
    Next
End Function

'*****************************************************************************
'[概要] Bitmap形式で画像を保存する
'[引数] ファイル名
'[戻値] True:成功
'*****************************************************************************
Public Function SaveToBmp(ByVal strFilename As String) As Boolean
On Error GoTo Finalization
    Dim BmpFileHeader As BITMAPFILEHEADER
    Dim BmpInfHeader  As BITMAPINFOHEADER
    
    With BmpFileHeader
        .Type = &H4D42 'BM'
        .OffBits = Len(BmpFileHeader) + Len(BmpInfHeader)
        .Size = .OffBits + FWidth * 4 * FHeight '32Bitの時
    End With
    
    With BmpInfHeader
        .Size = Len(BmpInfHeader)
        .Width = FWidth
        .Height = FHeight
        .Planes = 1 '常に1
        .BitCount = 32 'Bit
        .Compression = 0 'BI_RGB(無圧縮)
        .SizeImage = 0 '省略可
    End With
    
    Dim File As Integer
    Dim blnOpen As Boolean
    File = FreeFile()
    Open strFilename For Binary Access Write As #File
    blnOpen = True
    Put #File, , BmpFileHeader
    Put #File, , BmpInfHeader
    '左下から上に向かって設定するため水平方向に反転させる
    Call FlipHorizontal
    Put #File, , FPixels()
    Call FlipHorizontal '反転を元に戻す
    
    SaveToBmp = True
Finalization:
    If blnOpen Then Close #File
End Function

'*****************************************************************************
'[概要] 参考
'*****************************************************************************
Public Function LoadToIcon(ByVal strFilename As String) As Boolean
    Dim IconFileHeader As ICONDIR
    Dim IconInfHeader  As ICONDIRENTRY
    Dim BmpInfHeader   As BITMAPINFOHEADER
    
    Dim File As Integer
    File = FreeFile()
    Open strFilename For Binary Access Read As #File
    Get #File, , IconFileHeader
    Get #File, , IconInfHeader
    Get #File, , BmpInfHeader
    Close #File
End Function

'*****************************************************************************
'[概要] バイナリファイルを読込む
'[引数] ファイル名、読込むセルの位置、1行当たり何バイト表示するか
'*****************************************************************************
Public Sub LoadBinaryFile(ByVal strFilename As String, ByRef objCell As Range, Optional ColSize As Long = 16)
On Error GoTo Finalization
    Dim RowSize As Long
    RowSize = Int((FileLen(strFilename) + ColSize - 1) / ColSize)
    ReDim Data(1 To ColSize, 1 To RowSize) As Byte
        
    Dim File As Integer
    Dim blnOpen As Boolean
    File = FreeFile()
    Open strFilename For Binary Access Read As #File
    Get #File, , Data
    Close #File

    Dim x As Long
    Dim y As Long
    For y = 1 To RowSize
        For x = 1 To ColSize
            objCell.Cells(y, x) = Data(x, y)
        Next
    Next
Finalization:
    If blnOpen Then Close #File
End Sub

'*****************************************************************************
'[概要] hBitmapからFPixelsにアルファチェネルありの色を設定する
'[引数] Bitmapハンドル
'[戻値] なし
'*****************************************************************************
Public Sub GetPixelsFromHBITMAP(ByVal hBitmap As LongPtr)
On Error GoTo Finalization
'    Dim Bmp    As TBITMAP
'    Call GetObject(CommandBars.GetImageMso("Paste", 32, 32).Handle, Len(Bmp), Bmp)
    
    Dim hDC   As LongPtr
    Dim BmpInfHeader   As BITMAPINFOHEADER
    hDC = CreateCompatibleDC(0)
    BmpInfHeader.Size = 40 '40固定
    Call GetDIBits(hDC, hBitmap, 0, 0, ByVal 0, BmpInfHeader, DIB_RGB_COLORS)
    Call SetSize(BmpInfHeader.Width, Abs(BmpInfHeader.Height))
    
    With BmpInfHeader
        .Height = -FHeight  'ボトムアップからトップダウンに変更
        .Planes = 1 '常に1
        .BitCount = 32 'Bit
        .Compression = 0 'BI_RGB(無圧縮)
        .SizeImage = 0 '省略可
    End With
    
    Call GetDIBits(hDC, hBitmap, 0, FHeight, FPixels(1, 1), BmpInfHeader, DIB_RGB_COLORS)
Finalization:
    If hDC <> 0 Then Call DeleteDC(hDC)
End Sub

'*****************************************************************************
'[概要] IPictureに変換する
'[引数] Bitmapハンドル
'[戻値] なし
'*****************************************************************************
Public Function SetToIPicture() As IPicture
    Dim uPicInfo As PICTDESC_BMP
    Dim hGdiImage As LongPtr
    Dim hBitmap As LongPtr
    Dim IBitmap As IPicture
On Error GoTo Finalization
    If GdipCreateBitmapFromScan0(FWidth, FHeight, 0, PixelFormat32bppARGB, ByVal 0&, hGdiImage) <> S_OK Then
        Exit Function
    End If
    Call SetPixelsToBitmap(hGdiImage)
    If GdipCreateHBITMAPFromBitmap(hGdiImage, hBitmap, 0) <> S_OK Then GoTo Finalization

    With uPicInfo
        .Size = Len(uPicInfo)
        .Type = PICTYPE_BITMAP
        .hBitmap = hBitmap
        .hPal = 0
    End With

    Call OleCreatePictureIndirect(uPicInfo, GetCLSID(IID_IPictureDisp), True, IBitmap)
    Set SetToIPicture = IBitmap
Finalization:
    If hGdiImage <> 0 Then Call GdipDisposeImage(hGdiImage)
End Function

'*****************************************************************************
'[概要] ARGB色が同じかどうか判定する
'[引数] 対象の色
'[戻値] True:一致
'*****************************************************************************
Private Function SameColor2(ByRef Color1 As TRGBQuad, ByRef Color2 As TRGBQuad) As Boolean
    SameColor2 = (Color1.Alpha = Color2.Alpha And _
                  Color1.Red = Color2.Red And _
                  Color1.Green = Color2.Green And _
                  Color1.Blue = Color2.Blue)
End Function

'*****************************************************************************
'[概要] セルの色が同じかどうか判定
'[引数] 対象のセル
'[戻値] True:同じ
'*****************************************************************************
Public Function SameCellColor(ByRef objCell1 As Range, ByRef objCell2 As Range) As Boolean
    SameCellColor = SameColor2(CellToColor(objCell1), CellToColor(objCell2))
    Debug.Print
    Debug.Print objCell1.Address(0, 0), CellToColor(objCell1).Red & CellToColor(objCell1).Green & CellToColor(objCell1).Blue
    Debug.Print objCell2.Address(0, 0), CellToColor(objCell2).Red & CellToColor(objCell2).Green & CellToColor(objCell2).Blue
End Function

'*****************************************************************************
'[概要] RGBおよびアルファ値を増減させる
'[引数] lngUp:増加値(マイナスは減少値)、RGBαのいずれを増減の対象とするかどうか
'[戻値] なし
'*****************************************************************************
Public Sub AdjustColor(ByVal lngUp As Long, ByVal blnRed As Boolean, ByVal blnGreen As Boolean, ByVal blnBlue As Boolean, ByVal blnAlpha As Boolean)
    Dim x As Long, y As Long
    For x = 1 To FWidth
        For y = 1 To FHeight
            If lngUp > 0 Then
                If blnRed Then
                    FPixels(x, y).Red = WorksheetFunction.Min(255, FPixels(x, y).Red + lngUp)
                End If
                If blnGreen Then
                    FPixels(x, y).Green = WorksheetFunction.Min(255, FPixels(x, y).Green + lngUp)
                End If
                If blnBlue Then
                    FPixels(x, y).Blue = WorksheetFunction.Min(255, FPixels(x, y).Blue + lngUp)
                End If
                If blnAlpha Then
                    FPixels(x, y).Alpha = WorksheetFunction.Min(255, FPixels(x, y).Alpha + lngUp)
                End If
            Else
                If blnRed Then
                    FPixels(x, y).Red = WorksheetFunction.Max(0, FPixels(x, y).Red + lngUp)
                End If
                If blnGreen Then
                    FPixels(x, y).Green = WorksheetFunction.Max(0, FPixels(x, y).Green + lngUp)
                End If
                If blnBlue Then
                    FPixels(x, y).Blue = WorksheetFunction.Max(0, FPixels(x, y).Blue + lngUp)
                End If
                If blnAlpha Then
                    FPixels(x, y).Alpha = WorksheetFunction.Max(0, FPixels(x, y).Alpha + lngUp)
                End If
            End If
        Next
    Next
End Sub

'*****************************************************************************
'[概要] 塗潰す
'[引数] x,y:塗潰し開始位置、objColorCell塗潰し色のセル
'[戻値] なし
'*****************************************************************************
Public Sub Fill(ByVal x As Long, ByVal y As Long, ByRef objColorCell As Range)
    Dim dstColor As TRGBQuad
    dstColor = CellToColor(objColorCell)
    If SameColor2(FPixels(x, y), dstColor) Then
        Exit Sub
    End If
    
    'ワーク用の2次元配列を確保(上下左右に壁用のマスを確保する)
    '格納する値(1:開始セルと同色のセル、2:塗潰し完了セル、0:対象外の色のセルまたは壁)
    ReDim Pixels(0 To FWidth + 1, 0 To FHeight + 1) As Byte
    Dim xx As Long
    Dim yy As Long
    For xx = 1 To FWidth
        For yy = 1 To FHeight
            If SameColor2(FPixels(xx, yy), FPixels(x, y)) Then
                '開始セルと同色のセルに1を設定
                Pixels(xx, yy) = 1
            End If
        Next
    Next

    '塗潰し対象セル全てに2を設定する
    Do While FillCell(x, y, Pixels()): Loop
    
    '塗潰し対象セルに該当色を設定
    For xx = 1 To FWidth
        For yy = 1 To FHeight
            If Pixels(xx, yy) = 2 Then
                FPixels(xx, yy) = dstColor
            End If
        Next
    Next
Finalization:
End Sub

'*****************************************************************************
'[概要] 1マスでもアルファチャンネルを使用しているか判定
'[引数] なし
'[戻値] True:アルファチャンネルあり
'*****************************************************************************
Private Function HasAlpha() As Boolean
    Dim x As Long, y As Long
    Dim Color As TOLE_COLOR
    For x = 1 To FWidth
        For y = 1 To FHeight
            If FPixels(x, y).Alpha <> 0 Then
                HasAlpha = True
                Exit Function
            End If
        Next
    Next
End Function

'*****************************************************************************
'[概要] 対象座標(x,y)を塗潰し、一筆書きの次の座標を(x,y)に設定する
'       一筆書きが行詰れば次の起点となる座標を(x,y)に設定する
'[引数] x,y:座標、Pixels():塗潰し用の配列
'[戻値] True：次に塗潰すセルあり
'*****************************************************************************
Private Function FillCell(ByRef x As Long, ByRef y As Long, ByRef Pixels() As Byte) As Boolean
    FillCell = True
    Pixels(x, y) = 2 '塗潰し
    
    '上に自分と同じ色があるか
    If Pixels(x, y - 1) = 1 Then
        y = y - 1
        Exit Function
    End If
    '下に自分と同じ色があるか
    If Pixels(x, y + 1) = 1 Then
        y = y + 1
        Exit Function
    End If
    '右に自分と同じ色があるか
    If Pixels(x + 1, y) = 1 Then
        x = x + 1
        Exit Function
    End If
    '左に自分と同じ色があるか
    If Pixels(x - 1, y) = 1 Then
        x = x - 1
        Exit Function
    End If
    
    '一筆書きが行詰った時、次の起点となる座標を(x,y)に設定する
    Dim xx As Long
    Dim yy As Long
    For xx = 1 To FWidth
        For yy = 1 To FHeight
            If Pixels(xx, yy) = 1 Then '開始セルと同じ色
                '隣接するセルにすでに塗潰したセルがあるか
                If Pixels(xx - 1, yy) = 2 Or _
                   Pixels(xx + 1, yy) = 2 Or _
                   Pixels(xx, yy - 1) = 2 Or _
                   Pixels(xx, yy + 1) = 2 Then
                    x = xx
                    y = yy
                    Exit Function
                End If
            End If
        Next
    Next
    
    '次に塗潰せるセルなし
    FillCell = False
End Function

'以下は64×64ではスタックオーバーフローのため却下
'Private Sub Fill2(ByVal x As Byte, ByVal y As Byte, ByRef Pixels() As Long)
'    Pixels(x, y) = 2 '塗潰し
'    '上に自分と同じ色があるか
'    If Pixels(x, y - 1) = 1 Then
'        Call Fill2(x, y - 1, Pixels()) 'あればその座標で再帰呼び出し
'    End If
'    '下に自分と同じ色があるか
'    If Pixels(x, y + 1) = 1 Then
'        Call Fill2(x, y + 1, Pixels())
'    End If
'    '右に自分と同じ色があるか
'    If Pixels(x + 1, y) = 1 Then
'        Call Fill2(x + 1, y, Pixels())
'    End If
'    '左に自分と同じ色があるか
'    If Pixels(x - 1, y) = 1 Then
'        Call Fill2(x - 1, y, Pixels())
'    End If
'End Sub

