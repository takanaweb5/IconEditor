VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CGdiplus"
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

Private Type BITMAPDATA
    Width       As Long
    Height      As Long
    stride      As Long
    PixelFormat As Long
    scan0       As LongPtr
    Reserved    As LongPtr
End Type

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
Private Declare PtrSafe Function GdipCreateBitmapFromGdiDib Lib "gdiplus" (ByRef gdiBitmapInfo As Any, ByRef gdiBitmapData As Any, bitmap As LongPtr) As Long
Private Declare PtrSafe Function GdipDrawImageRectI Lib "gdiplus.dll" (ByVal GpGraphics As LongPtr, ByVal GpImage As LongPtr, ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long) As Long
Private Declare PtrSafe Function GdipSetInterpolationMode Lib "gdiplus" (ByVal graphics As LongPtr, ByVal nInterpolationMode As Long) As Long
Private Declare PtrSafe Function GdipCreateBitmapFromGraphics Lib "gdiplus" (ByVal Width As Long, ByVal Height As Long, ByVal Target As LongPtr, bitmap As LongPtr) As Long
Private Declare PtrSafe Function GdipSetPixelOffsetMode Lib "gdiplus" (ByVal graphics As LongPtr, ByVal PixOffsetMode As Long) As Long
Private Declare PtrSafe Function GdipBitmapLockBits Lib "gdiplus" (ByVal bitmap As LongPtr, ByRef RECT As Any, ByVal flags As Long, ByVal PixelFormat As Long, ByRef lockedBitmapData As Any) As Long
Private Declare PtrSafe Function GdipBitmapUnlockBits Lib "gdiplus" (ByVal bitmap As LongPtr, ByRef lockedBitmapData As Any) As Long
Private Declare PtrSafe Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As LongPtr, ByRef graphics As LongPtr) As Long
Private Declare PtrSafe Function CLSIDFromString Lib "ole32" (ByVal lpszCLSID As LongPtr, ByRef pclsid As Guid) As Long

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

Private Const CLSID_Suffix As String = "-1A04-11D3-9A73-0000F81EF32E}"
Private Const CLSID_BMP    As String = "{557CF400" & CLSID_Suffix
Private Const CLSID_JPG    As String = "{557CF401" & CLSID_Suffix
Private Const CLSID_GIF    As String = "{557CF402" & CLSID_Suffix
Private Const CLSID_EMF    As String = "{557CF403" & CLSID_Suffix
Private Const CLSID_WMF    As String = "{557CF404" & CLSID_Suffix
Private Const CLSID_TIF    As String = "{557CF405" & CLSID_Suffix
Private Const CLSID_PNG    As String = "{557CF406" & CLSID_Suffix
Private Const CLSID_ICO    As String = "{557CF407" & CLSID_Suffix
Private Const PixelOffsetModeHalf = 4
Private Const PixelFormat32bppARGB = &H26200A
Private Const S_OK = 0

Private FGdiPlus As LongPtr
Private FGdiImage As LongPtr

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
    If FGdiPlus <> 0 Then
        Call GdiplusShutdown(FGdiPlus)
    End If
End Sub

'*****************************************************************************
'[概要] 画像の幅
'[引数] なし
'*****************************************************************************
Public Property Get Width() As Long
    Dim lngWidth As Long
    Call GdipGetImageWidth(FGdiImage, lngWidth)
    Width = lngWidth
End Property

'*****************************************************************************
'[概要] 画像の高さ
'[引数] なし
'*****************************************************************************
Public Property Get Height() As Long
    Dim lngHeight As Long
    Call GdipGetImageHeight(FGdiImage, lngHeight)
    Height = lngHeight
End Property

'*****************************************************************************
'[概要] ファイルからGdiImageオブジェクトを作成
'[引数] ファイル名
'[戻値] なし
'*****************************************************************************
Public Sub CreateFromFile(ByVal strFile As String)
    If GdipLoadImageFromFile(ByVal StrPtr(strFile), FGdiImage) <> S_OK Then
        Call Err.Raise(513, "GdipLoadImageFromFileエラー")
    End If
    Call GdipBitmapSetResolution(FGdiImage, 96#, 96#) '標準のDPIに設定
End Sub

'*****************************************************************************
'[概要] StreamオブジェクトからGdiImageオブジェクトを作成
'[引数] ファイル名
'[戻値] なし
'*****************************************************************************
Public Sub CreateFromStream(ByVal Stream As IUnknown)
    If GdipLoadImageFromStream(ObjPtr(Stream), FGdiImage) <> S_OK Then
        Call Err.Raise(513, "GdipLoadImageFromStreamエラー")
    End If
    Call GdipBitmapSetResolution(FGdiImage, 96#, 96#) '標準のDPIに設定
End Sub

'*****************************************************************************
'※アルファチャネルが反映されないため却下
'[概要] ビットマップハンドルからGdiImageオブジェクトを作成
'[引数] ファイル名
'[戻値] なし
'*****************************************************************************
'Public Sub CreateFromHBITMAP(ByVal hBitmap As LongPtr)
'    If GdipCreateBitmapFromHBITMAP(hBitmap, 0&, FGdiImage) <> S_OK Then
'        Call Err.Raise(513, "GdipCreateBitmapFromHBITMAPエラー")
'    End If
'    Call GdipBitmapSetResolution(FGdiImage, 96#, 96#) '標準のDPIに設定
'End Sub

'*****************************************************************************
'[概要] Pixel情報からGdiImageオブジェクトを作成
'[引数] ファイル名
'[戻値] なし
'*****************************************************************************
Public Sub CreateFromPixels(ByRef Pixels() As Long)
    Dim lngWidth  As Long
    Dim lngHeight As Long
    lngWidth = UBound(Pixels, 1)
    lngHeight = UBound(Pixels, 2)
    
    'GdiImageオブジェクトを作成
    If GdipCreateBitmapFromScan0(lngWidth, lngHeight, lngWidth * 4, PixelFormat32bppARGB, Pixels(1, 1), FGdiImage) <> S_OK Then
        Call Err.Raise(513, "GdipCreateBitmapFromScan0エラー")
    End If
    Call GdipBitmapSetResolution(FGdiImage, 96#, 96#) '標準のDPIに設定
    
'    Dim BmpData As BITMAPDATA
'    With BmpData
'        .Width = lngWidth
'        .Height = lngHeight
'        .PixelFormat = PixelFormat32bppARGB
'        .scan0 = VarPtr(Pixels(1, 1))
'        .stride = lngWidth * 4
'    End With
'    If GdipBitmapLockBits(FGdiImage, ByVal 0, ImageLockModeUserInputBuf Or ImageLockModeWrite, PixelFormat32bppARGB, BmpData) <> S_OK Then
'        Call Err.Raise(513, "GdipBitmapLockBitsエラー")
'    End If
'    Call GdipBitmapUnlockBits(FGdiImage, BmpData)

'    Dim x As Long, y As Long
'    For x = 1 To lngWidth
'        For y = 1 To lngHeight
'            Call GdipBitmapSetPixel(FGdiImage, x - 1, y - 1, CastARGB(Pixels(x, y)))
'        Next
'    Next
End Sub

'*****************************************************************************
'[概要] Pixel情報からGdiImageオブジェクトを作成
'[引数] ファイル名
'[戻値] なし
'*****************************************************************************
Public Sub ToPixels(ByRef Pixels() As Long)
    Dim lngWidth  As Long
    Dim lngHeight As Long
    lngWidth = Width
    lngHeight = Height
    ReDim Pixels(1 To Width, 1 To Height)
    
    Dim BmpData As BITMAPDATA
    With BmpData
        .Width = lngWidth
        .Height = lngHeight
        .PixelFormat = PixelFormat32bppARGB
        .scan0 = VarPtr(Pixels(1, 1))
        .stride = lngWidth * 4
    End With
    
    If GdipBitmapLockBits(FGdiImage, ByVal 0, ImageLockModeUserInputBuf Or ImageLockModeRead, PixelFormat32bppARGB, BmpData) <> S_OK Then
        Call Err.Raise(513, "GdipBitmapLockBitsエラー")
    End If
    Call GdipBitmapUnlockBits(FGdiImage, BmpData)
End Sub

'*****************************************************************************
'[概要] GdiImageからビットマップハンドルを作成
'[引数] なし
'[戻値] ビットマップハンドル
'*****************************************************************************
Public Function ToHBITMAP() As LongPtr
    Dim hBitmap As LongPtr
    If GdipCreateHBITMAPFromBitmap(FGdiImage, hBitmap, 0) <> S_OK Then
        Call Err.Raise(513, "GdipCreateHBITMAPFromBitmapエラー")
    End If
    ToHBITMAP = hBitmap
End Function

'*****************************************************************************
'[概要] GdiImageからPNGのファイルStreamを作成
'[引数] 戻値：PNGのファイルStream
'[戻値] なし
'*****************************************************************************
Public Sub ToPNGStream(ByRef Stream As IUnknown)
    Dim gGuid As Guid
    Call CLSIDFromString(StrPtr(CLSID_PNG), gGuid)
    If GdipSaveImageToStream(FGdiImage, ByVal ObjPtr(Stream), gGuid, 0) <> S_OK Then
        Call Err.Raise(513, "GdipSaveImageToStreamエラー")
    End If
End Sub

'*****************************************************************************
'[概要] PNGファイルで保存する
'[引数] ファイル名
'[戻値] なし
'*****************************************************************************
Public Sub SaveToPNGFile(ByVal strFile As String)
    Dim gGuid As Guid
    Call CLSIDFromString(StrPtr(CLSID_PNG), gGuid)
    If GdipSaveImageToFile(FGdiImage, ByVal StrPtr(strFile), gGuid, 0) <> S_OK Then
        Call Err.Raise(513, "GdipSaveImageToFileエラー")
    End If
End Sub

'*****************************************************************************
'[概要] GdiImageオブジェクトの画像のサイズを変更する
'[引数] 新しいサイズ
'[戻値] なし
'*****************************************************************************
Public Sub Resize(ByVal lngWidth As Long, ByVal lngHeight As Long)
On Error GoTo Finalization
    '新しいサイズのBitmapオブジェクトを作成
    Dim hNewBmp  As LongPtr
    If GdipCreateBitmapFromScan0(lngWidth, lngHeight, 0, PixelFormat32bppARGB, ByVal 0&, hNewBmp) <> S_OK Then
        Call Err.Raise(513, "GdipCreateBitmapFromScan0エラー")
    End If
 
    'サイズを変更
    Dim hGraphics As LongPtr
    If GdipGetImageGraphicsContext(hNewBmp, hGraphics) = 0 Then
'       Call GdipSetInterpolationMode(hGraphics, InterpolationModeHighQuality)
        Call GdipSetInterpolationMode(hGraphics, InterpolationModeNearestNeighbor)
        Call GdipSetPixelOffsetMode(hGraphics, PixelOffsetModeHalf)
        '新しいサイズで描画
        Call GdipDrawImageRectI(hGraphics, FGdiImage, 0, 0, lngWidth, lngHeight)
    End If
    
    '変更前のGdiImageオブジェクトを解放
    Call GdipDisposeImage(FGdiImage)
    '変更後のGdiImageオブジェクトを割当て
    FGdiImage = hNewBmp
Finalization:
    If hGraphics <> 0 Then Call GdipDeleteGraphics(hGraphics)
    If Err.Number <> 0 Then Call Err.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

