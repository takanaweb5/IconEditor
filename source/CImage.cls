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

'Private Enum GpStatus
'   Ok = 0
'   GenericError = 1
'   InvalidParameter = 2
'   OutOfMemory = 3
'   ObjectBusy = 4
'   InsufficientBuffer = 5
'   NotImplemented = 6
'   Win32Error = 7
'   WrongState = 8
'   Aborted = 9
'   FileNotFound = 10
'   ValueOverflow = 11
'   AccessDenied = 12
'   UnknownImageFormat = 13
'   FontFamilyNotFound = 14
'   FontStyleNotFound = 15
'   NotTrueTypeFont = 16
'   UnsupportedGdiplusVersion = 17
'   GdiplusNotInitialized = 18
'   PropertyNotFound = 19
'   PropertyNotSupported = 20
'End Enum

Private Type Guid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Private Type PICTDESC_ICON
    size   As Long
    Type   As Long
    hIcon  As LongPtr
End Type

Private Type PICTDESC_BMP
    size    As Long
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
    size      As Long
    Reserved1 As Integer
    Reserved2 As Integer
    OffBits   As Long
End Type

Private Type BITMAPINFOHEADER
    size          As Long
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
Private Declare PtrSafe Function GdipBitmapLockBits Lib "gdiplus" (ByVal bitmap As LongPtr, ByRef RECT As Any, ByVal flags As Long, ByVal PixelFormat As Long, ByRef lockedBitmapData As Any) As Long
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
Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long

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

'*****************************************************************************
'[概要] コンストラクタ
'*****************************************************************************
'Private Sub Class_Initialize()
'    Dim GdiInput As GdiplusStartupInput
'    GdiInput.GdiplusVersion = 1&
'    If GdiplusStartup(FGdiPlus, GdiInput) <> S_OK Then
'        Call Err.Raise(513, "GdiplusStartupエラー")
'    End If
'
'    Call MsgBox("Class_Initialize")
'End Sub

'*****************************************************************************
'[概要] デストラクタ
'*****************************************************************************
'Private Sub Class_Terminate()
'    If FGdiPlus <> 0 Then Call GdiplusShutdown(FGdiPlus)
'    Call MsgBox("Class_Terminate")
'End Sub

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
Private Sub Transpose()
    ReDim work(1 To FHeight, 1 To FWidth) As TRGBQuad
    Dim x As Long, y As Long
    For x = 1 To FHeight
        For y = 1 To FWidth
            work(x, y) = FPixels(y, x)
        Next
    Next
    Dim Swap As Long
    Swap = FWidth
    FWidth = FHeight
    FHeight = Swap
    FPixels = work
End Sub

'*****************************************************************************
'[概要] 水平方向に画像を反転させる
'[引数] なし
'[戻値] なし
'*****************************************************************************
Public Sub FlipHorizontal()
    ReDim work(1 To FWidth, 1 To FHeight) As TRGBQuad
    Dim x As Long, y As Long
    For x = 1 To FWidth
        For y = 1 To FHeight
            work(x, y) = FPixels(x, FHeight - y + 1)
        Next
    Next
    FPixels = work
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
Public Sub LoadImageFromFile(ByVal strFile As String)
On Error GoTo Finalization
    Dim hGdiImage As LongPtr
    If GdipLoadImageFromFile(ByVal StrPtr(strFile), hGdiImage) <> S_OK Then
        Call Err.Raise(513, "GdipLoadImageFromFileエラー")
    End If
    Call GdipBitmapSetResolution(hGdiImage, 96#, 96#) '標準のDPIに設定
    Call GdipGetImageWidth(hGdiImage, FWidth)
    Call GdipGetImageHeight(hGdiImage, FHeight)
    ReDim FPixels(1 To FWidth, 1 To FHeight)
    Call GetPixelsFromImage(hGdiImage)
Finalization:
    If hGdiImage <> 0 Then Call GdipDisposeImage(hGdiImage)
    If Err.Number <> 0 Then Call Err.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'*****************************************************************************
'[概要] セルのデータからアイコンファイルを読込む
'[引数] データを取得する行(Rangeオブジェクト)
'[戻値] なし
'*****************************************************************************
Public Sub LoadImageFromResource(ByRef objRow As Range)
On Error GoTo Finalization
    'ファイルサイズの配列を作成
    ReDim Data(1 To objRow.Cells(1, 1).End(xlToRight).Column - 1) As Byte
    Dim x As Long
    For x = 1 To UBound(Data)
         Data(x) = objRow.Cells(1, x + 1)
    Next
    
    Dim Stream    As IUnknown
    Dim hGdiImage As LongPtr
    If CreateStreamOnHGlobal(VarPtr(Data(1)), 0, Stream) <> S_OK Then GoTo Finalization
    
    If GdipLoadImageFromStream(ObjPtr(Stream), hGdiImage) <> S_OK Then
        Call Err.Raise(513, "LoadImageFromResourceエラー")
    End If
    Call GdipBitmapSetResolution(hGdiImage, 96#, 96#) '標準のDPIに設定
    Call GdipGetImageWidth(hGdiImage, FWidth)
    Call GdipGetImageHeight(hGdiImage, FHeight)
    ReDim FPixels(1 To FWidth, 1 To FHeight)
    Call GetPixelsFromImage(hGdiImage)
Finalization:
    If hGdiImage <> 0 Then Call GdipDisposeImage(hGdiImage)
    If Err.Number <> 0 Then Call Err.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'*****************************************************************************
'[概要] ファイルに画像を保存する
'[引数] ファイル名
'[戻値] なし
'*****************************************************************************
Public Sub SaveImageToFile(ByVal strFile As String)
On Error GoTo Finalization
    Dim strGuid As String
    Select Case GetFileExtension(strFile)
    Case "ICO"
        Call SaveToIcon(strFile) 'アルファチェネルの設定あり
        Exit Sub
    Case "BMP"
        Call SaveToBmp(strFile) 'アルファチェネルの設定あり
        Exit Sub
    Case "PNG"
        strGuid = CLSID_PNG
    Case Else
        strGuid = CLSID_BMP
    End Select
    
    Dim hGdiImage As LongPtr
    If GdipCreateBitmapFromScan0(FWidth, FHeight, 0, PixelFormat32bppARGB, ByVal 0&, hGdiImage) <> S_OK Then
        Call Err.Raise(513, "GdipCreateBitmapFromScan0エラー")
    End If
    Call SetPixelsToImage(hGdiImage)
    Call GdipBitmapSetResolution(hGdiImage, 96#, 96#) '標準のDPIに設定
    If GdipSaveImageToFile(hGdiImage, ByVal StrPtr(strFile), GetCLSID(strGuid), 0) <> S_OK Then
        Call Err.Raise(513, "GdipSaveImageToFileエラー")
    End If
    Call GdipDisposeImage(hGdiImage)
    Exit Sub
Finalization:
    If hGdiImage <> 0 Then Call GdipDisposeImage(hGdiImage)
    If Err.Number <> 0 Then Call Err.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
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
        Exit Sub
    End If

On Error GoTo Finalization
    '元のサイズのBitmapオブジェクトを作成
    If GdipCreateBitmapFromScan0(FWidth, FHeight, 0, PixelFormat32bppARGB, ByVal 0&, hOldBmp) <> S_OK Then
        Call Err.Raise(513, "GdipCreateBitmapFromScan0エラー")
    End If
    Call SetPixelsToImage(hOldBmp)
    
'    Dim hSrcBitmap As LongPtr
'    Call GdipCreateHBITMAPFromBitmap(hOldBmp, hSrcBitmap, 0)
'    Dim hSrcDC As LongPtr
'    hSrcDC = CreateCompatibleDC(0)
'    Call SelectObject(hSrcDC, hSrcBitmap)
'
'    Dim hDstBitmap As LongPtr
'    hDstBitmap = CreateBitmap(lngWidth, lngHeight, 1, 32, ByVal 0&)
'    Dim hDstDc As LongPtr
'    hDstDc = CreateCompatibleDC(0)
'    Call SelectObject(hDstDc, hDstBitmap)
'
'    Call StretchBlt(hDstDc, 0, 0, lngWidth, lngHeight, hSrcDC, 0, 0, FWidth, FHeight, SRCCOPY)
'    Call SetSize(lngWidth, lngHeight)
'    Call GetPixelsFromHBITMAP(hDstBitmap)
    
    '新しいサイズのBitmapオブジェクトを作成
    If GdipCreateBitmapFromScan0(lngWidth, lngHeight, 0, PixelFormat32bppARGB, ByVal 0&, hNewBmp) <> S_OK Then
        Call Err.Raise(513, "GdipCreateBitmapFromScan0エラー")
    End If
 
    'サイズを変更
    If GdipGetImageGraphicsContext(hNewBmp, hNewGraphics) = 0 Then
'        Call GdipSetInterpolationMode(hGraphics, InterpolationModeHighQuality)
        Call GdipSetInterpolationMode(hNewGraphics, InterpolationModeNearestNeighbor)
        Call GdipSetPixelOffsetMode(hNewGraphics, PixelOffsetModeHalf)
        Call GdipDrawImageRectI(hNewGraphics, hOldBmp, 0, 0, lngWidth, lngHeight)
    End If
    
    Call SetSize(lngWidth, lngHeight)
    Call GetPixelsFromImage(hNewBmp)
Finalization:
    If hNewGraphics <> 0 Then Call GdipDeleteGraphics(hNewGraphics)
    If hNewBmp <> 0 Then Call GdipDisposeImage(hNewBmp)
    If hOldBmp <> 0 Then Call GdipDisposeImage(hOldBmp)
    If Err.Number <> 0 Then Call Err.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'*****************************************************************************
'[概要] BitmapオブジェクトからFPixels()を設定
'[引数] Bitmapオブジェクト
'[戻値] なし
'*****************************************************************************
Private Sub GetPixelsFromImage(ByVal hGdiImage As LongPtr)
    Dim BmpData As BITMAPDATA
    With BmpData
        .Width = FWidth
        .Height = FHeight
        .PixelFormat = PixelFormat32bppARGB
        .scan0 = VarPtr(FPixels(1, 1))
        .stride = FWidth * 4
    End With
    
    If GdipBitmapLockBits(hGdiImage, ByVal 0, ImageLockModeUserInputBuf Or ImageLockModeRead, PixelFormat32bppARGB, BmpData) <> S_OK Then
        Call Err.Raise(513, "GdipBitmapLockBitsエラー")
    End If
    Call GdipBitmapUnlockBits(hGdiImage, BmpData)
    
    'αチャンネルがすべて初期値の時は、全てのPixelのα値を255に再設定する
    Call SetNoAlpha
End Sub

'*****************************************************************************
'[概要] FPixels()からBitmapオブジェクトを設定 ※アルファチャネルあり
'[引数] Bitmapオブジェクト
'[戻値] なし
'*****************************************************************************
Private Sub SetPixelsToImage(ByRef hGdiImage As LongPtr)
    Dim BmpData As BITMAPDATA
    With BmpData
        .Width = FWidth
        .Height = FHeight
        .PixelFormat = PixelFormat32bppARGB
        .scan0 = VarPtr(FPixels(1, 1))
        .stride = FWidth * 4
    End With

    If GdipBitmapLockBits(hGdiImage, ByVal 0, ImageLockModeUserInputBuf Or ImageLockModeWrite, PixelFormat32bppARGB, BmpData) <> S_OK Then
        Call Err.Raise(513, "GdipBitmapLockBitsエラー")
    End If
    Call GdipBitmapUnlockBits(hGdiImage, BmpData)

'    Dim x As Long, y As Long
'    For x = 1 To FWidth
'        For y = 1 To FHeight
'            Call GdipBitmapSetPixel(hGdiImage, x - 1, y - 1, CastARGB(FPixels(x, y)))
'        Next
'    Next
End Sub

'*****************************************************************************
'[概要] BMP形式の画像をクリップボードから読込む
'[引数] なし
'[戻値] なし
'*****************************************************************************
Public Sub LoadBMPFromClipbord()
    If OpenClipboard(0) = 0 Then Exit Sub
On Error GoTo Finalization
    Dim hBitmap As LongPtr
    hBitmap = GetClipboardData(CF_BITMAP)
    If hBitmap = 0 Then GoTo Finalization
    Call GetPixelsFromHBITMAP(hBitmap)
Finalization:
    Call CloseClipboard
End Sub

'*****************************************************************************
'[概要] PNG形式の画像をクリップボードから読込む
'[引数] なし
'[戻値] なし
'*****************************************************************************
Public Sub LoadPNGFromClipbord()
    If OpenClipboard(0) = 0 Then Exit Sub
On Error GoTo Finalization
    Dim CF_PNG  As Long
    Dim hGlobal As LongPtr
    CF_PNG = RegisterClipboardFormat("PNG")
    hGlobal = GetClipboardData(CF_PNG)
    If hGlobal = 0 Then GoTo Finalization

    Dim Stream    As IUnknown
    Dim hGdiImage As LongPtr
    If CreateStreamOnHGlobal(ByVal hGlobal, 0, Stream) <> S_OK Then GoTo Finalization
    If GdipLoadImageFromStream(ObjPtr(Stream), hGdiImage) <> S_OK Then
        Call Err.Raise(513, "GdipLoadImageFromStreamエラー")
    End If
    Call GdipBitmapSetResolution(hGdiImage, 96#, 96#) '標準のDPIに設定
    Call GdipGetImageWidth(hGdiImage, FWidth)
    Call GdipGetImageHeight(hGdiImage, FHeight)
    ReDim FPixels(1 To FWidth, 1 To FHeight)
    Call GetPixelsFromImage(hGdiImage)
Finalization:
    Call CloseClipboard
    If hGdiImage <> 0 Then Call GdipDisposeImage(hGdiImage)
    If Err.Number <> 0 Then Call Err.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'*****************************************************************************
'[概要] PNG形式とBMP形式の画像をクリップボードに書込む
'[引数] なし
'[戻値] True:成功
'*****************************************************************************
Public Sub SaveImageToClipbord()
    If OpenClipboard(0) = 0 Then Exit Sub
On Error GoTo Finalization
    Call EmptyClipboard
    Call SavePNGToClipbord
    Call SaveBMPToClipbord
Finalization:
    Call CloseClipboard
End Sub

'*****************************************************************************
'[概要] PNG形式の画像をクリップボードに書込む
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub SavePNGToClipbord()
On Error GoTo ErrHandle
    Dim Stream    As IUnknown
    Dim hGdiImage As LongPtr
    If CreateStreamOnHGlobal(0, 1, Stream) <> S_OK Then GoTo ErrHandle
    '32Bitカラーのbitmapオブジェクトを生成
    If GdipCreateBitmapFromScan0(FWidth, FHeight, 0, PixelFormat32bppARGB, ByVal 0&, hGdiImage) <> S_OK Then
        Call Err.Raise(513, "GdipCreateBitmapFromScan0エラー")
    End If
    Call SetPixelsToImage(hGdiImage)
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
ErrHandle:
    If hGdiImage <> 0 Then Call GdipDisposeImage(hGdiImage)
    If Err.Number <> 0 Then Call Err.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'*****************************************************************************
'[概要] BMP形式の画像をクリップボードに書込む
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub SaveBMPToClipbord()
On Error GoTo Finalization
    Dim hGdiImage As LongPtr
    Dim hBitmap As LongPtr
    If GdipCreateBitmapFromScan0(FWidth, FHeight, 0, PixelFormat32bppARGB, ByVal 0&, hGdiImage) <> S_OK Then
        Exit Sub
    End If
    Call SetPixelsToImage(hGdiImage)
    If GdipCreateHBITMAPFromBitmap(hGdiImage, hBitmap, 0) <> S_OK Then GoTo Finalization
    If SetClipboardData(CF_BITMAP, hBitmap) = 0 Then GoTo Finalization
Finalization:
    If hGdiImage <> 0 Then Call GdipDisposeImage(hGdiImage)
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
'[引数] 色を設定するセル
'[戻値] なし
'*****************************************************************************
Public Sub SetPixelsToRange(ByVal objRange As Range)
    Dim objCell As Range
    Dim Color  As TRGBQuad
    
    '高速化のためいったん書式をクリア
    Call ClearRange(objRange.Resize(FHeight, FWidth))
    Dim x As Long, y As Long
    For x = 1 To FWidth
        For y = 1 To FHeight
            Call ColorToCell(objRange.Cells(y, x), FPixels(x, y), False)
        Next
    Next
End Sub

'*****************************************************************************
'[概要] アイコンファイルとして画像を保存する
'       GDI+では、減色されてアルファ情報も失われるため使用しない
'[引数] ファイル名
'[戻値] なし
'*****************************************************************************
Public Sub SaveToIcon(ByVal strFilename As String)
On Error GoTo Finalization
    Dim MaskSize    As Long
    Dim MaskBits()  As Byte
    
    '左下から上に向かって設定するため水平方向に反転させる
    Call FlipHorizontal
    MaskSize = MakeMaskBits(MaskBits())
    Call FlipHorizontal '反転を元に戻す
    
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
        .size = Len(BmpInfHeader)
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
'            FPixels(x, y).Alpha = 0 '検証用
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
    '左下から上に向かって設定するため水平方向に反転させる
    Call FlipHorizontal
    Put #File, , FPixels()
    Put #File, , MaskBits
    Call FlipHorizontal '反転を元に戻す
Finalization:
    If blnOpen Then Close #File
End Sub

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
'[戻値] なし
'*****************************************************************************
Public Sub SaveToBmp(ByVal strFilename As String)
On Error GoTo Finalization
    Dim BmpFileHeader As BITMAPFILEHEADER
    Dim BmpInfHeader  As BITMAPINFOHEADER
    
    With BmpFileHeader
        .Type = &H4D42 'BM'
        .OffBits = Len(BmpFileHeader) + Len(BmpInfHeader)
        .size = .OffBits + FWidth * 4 * FHeight '32Bitの時
    End With
    
    With BmpInfHeader
        .size = Len(BmpInfHeader)
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
Finalization:
    If blnOpen Then Close #File
End Sub

'*****************************************************************************
'[概要] hBitmapからFPixelsにアルファチャネルありの色を設定する
'[引数] Bitmapハンドル
'[戻値] なし
'*****************************************************************************
Public Sub GetPixelsFromHBITMAP(ByVal hBitmap As LongPtr)
On Error GoTo Finalization
    Dim hDC   As LongPtr
    Dim BmpInfHeader   As BITMAPINFOHEADER
    hDC = CreateCompatibleDC(0)
    BmpInfHeader.size = 40 '40固定
    
    '幅と高さのみ取得
    Call GetDIBits(hDC, hBitmap, 0, 0, ByVal 0, BmpInfHeader, DIB_RGB_COLORS)
    Call SetSize(BmpInfHeader.Width, Abs(BmpInfHeader.Height))
    
    With BmpInfHeader
        .Height = -FHeight  'ボトムアップからトップダウンに変更
        .Planes = 1 '常に1
        .BitCount = 32 'Bit
        .Compression = 0 'BI_RGB(無圧縮)
        .SizeImage = 0 '省略可
    End With
    
    'FPixels()に値を設定
    Call GetDIBits(hDC, hBitmap, 0, FHeight, FPixels(1, 1), BmpInfHeader, DIB_RGB_COLORS)
    
    'αチャンネルがすべて初期値の時は、全てのPixelのα値を255に再設定する
    Call SetNoAlpha
Finalization:
    If hDC <> 0 Then Call DeleteDC(hDC)
End Sub

'*****************************************************************************
'※アルファチャネルが反映されないため却下
'*****************************************************************************
'Public Sub GetPixelsFromHBITMAP(ByVal hBitmap As LongPtr)
'On Error GoTo Finalization
'    Dim hGdiImage As LongPtr
'    If GdipCreateBitmapFromHBITMAP(hBitmap, 0&, hGdiImage) <> S_OK Then
'        Call Err.Raise(513, "GdipCreateBitmapFromHBITMAPエラー")
'    End If
'    Call GdipGetImageWidth(hGdiImage, FWidth)
'    Call GdipGetImageHeight(hGdiImage, FHeight)
'    ReDim FPixels(1 To FWidth, 1 To FHeight)
'    Call GetPixelsFromImage(hGdiImage)
'Finalization:
'    If hGdiImage <> 0 Then Call GdipDisposeImage(hGdiImage)
'    If Err.Number <> 0 Then Call Err.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
'End Sub

'*****************************************************************************
'[概要] IPictureに変換する
'[引数] Bitmapハンドル
'[戻値] なし
'*****************************************************************************
Public Function SetToIPicture() As IPicture
On Error GoTo Finalization
    Dim uPicInfo As PICTDESC_BMP
    Dim hGdiImage As LongPtr
    Dim hBitmap As LongPtr
    Dim IBitmap As IPicture
    If GdipCreateBitmapFromScan0(FWidth, FHeight, 0, PixelFormat32bppARGB, ByVal 0&, hGdiImage) <> S_OK Then
        Exit Function
    End If
    Call SetPixelsToImage(hGdiImage)
    If GdipCreateHBITMAPFromBitmap(hGdiImage, hBitmap, 0) <> S_OK Then GoTo Finalization

    With uPicInfo
        .size = Len(uPicInfo)
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
'[概要] 塗潰す
'[引数] x,y:塗潰し開始位置、objColorCell塗潰し色のセル
'[戻値] なし
'*****************************************************************************
Public Sub Fill(ByVal x As Long, ByVal y As Long, ByRef objColorCell As Range)
    Dim DstColor As TRGBQuad
    DstColor = CellToColor(objColorCell)
    If SameColor(FPixels(x, y), DstColor) Then
        Exit Sub
    End If
    
    'ワーク用の2次元配列を確保(上下左右に壁用のマスを確保する)
    '格納する値(1:開始セルと同色のセル、2:塗潰し完了セル、0:対象外の色のセルまたは壁)
    ReDim Pixels(0 To FWidth + 1, 0 To FHeight + 1) As Byte
    Dim xx As Long
    Dim yy As Long
    For xx = 1 To FWidth
        For yy = 1 To FHeight
            If SameColor(FPixels(xx, yy), FPixels(x, y)) Then
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
                FPixels(xx, yy) = DstColor
            End If
        Next
    Next
Finalization:
End Sub

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

'*****************************************************************************
'[概要] アルファチャンネルを使用しているマスがない時、すべてのα値を255にする
'       ※放置すると透明な画像となってしまうため
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub SetNoAlpha()
    If HasAlpha Then
        Exit Sub
    End If
    
    Dim x As Long, y As Long
    For x = 1 To FWidth
        For y = 1 To FHeight
            FPixels(x, y).Alpha = 255 '不透明
        Next
    Next
End Sub

'*****************************************************************************
'[概要] 1マスでもアルファチャンネルを使用しているか判定
'[引数] なし
'[戻値] True:アルファチャンネルあり
'*****************************************************************************
Private Function HasAlpha() As Boolean
    Dim x As Long, y As Long
    For x = 1 To FWidth
        For y = 1 To FHeight
            If FPixels(x, y).Alpha <> 0 Then
                HasAlpha = True
                Exit Function
            End If
        Next
    Next
End Function

