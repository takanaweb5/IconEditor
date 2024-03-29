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

Private Type PICTDESC_ICON
    size   As Long
    Type   As Long
    hIcon  As LongPtr
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

Private Declare PtrSafe Function IIDFromString Lib "ole32" (ByVal lpsz As LongPtr, ByRef lpiid As Any) As Long
Private Declare PtrSafe Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As LongPtr, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Private Declare PtrSafe Function GetHGlobalFromStream Lib "ole32" (ByVal Stream As IUnknown, ByRef hGlobal As LongPtr) As Long
Private Declare PtrSafe Function IStream_Size Lib "shlwapi" Alias "#214" (ByVal Stream As IUnknown, ByRef pui As Currency) As Long

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

Private Const IID_IPictureDisp As String = "{7BF80981-BF32-101A-8BBB-00AA00300CAB}"

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
Private FPixels() As Long
Private FWidth As Long
Private FHeight As Long

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
    ReDim work(1 To FHeight, 1 To FWidth) As Long
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
    ReDim work(1 To FWidth, 1 To FHeight) As Long
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
'[概要] ファイルから画像を読込む
'[引数] ファイル名
'[戻値] なし
'*****************************************************************************
Public Sub LoadImageFromFile(ByVal strFile As String, Optional lngWidth As Long = 0, Optional lngHeight As Long = 0)
    Dim Gdip As New CGdiplus
    Call Gdip.CreateFromFile(strFile, lngWidth, lngHeight)
    Call Me.SetSize(Gdip.Width, Gdip.Height)
    Call Gdip.ToPixels(FPixels())
    Call SetNoAlpha
End Sub

'*****************************************************************************
'[概要] セルのデータからアイコンファイルを読込む
'[引数] データを取得する行(Rangeオブジェクト)
'[戻値] なし
'*****************************************************************************
Public Sub LoadImageFromResource(ByRef objRow As Range)
    'ファイルサイズの配列を作成
    ReDim Data(1 To objRow.Cells(1, 1).End(xlToRight).Column - 1) As Byte
    Dim x As Long
    For x = 1 To UBound(Data)
         Data(x) = objRow.Cells(1, x + 1)
    Next
    
    Dim Gdip As New CGdiplus
    Call Gdip.CreateFromHGlobal(VarPtr(Data(1)))
    Call Me.SetSize(Gdip.Width, Gdip.Height)
    Call Gdip.ToPixels(FPixels())
    Call SetNoAlpha
End Sub

'*****************************************************************************
'[概要] ファイルに画像を保存する
'[引数] ファイル名
'[戻値] なし
'*****************************************************************************
Public Sub SaveImageToFile(ByVal strFile As String)
    Select Case GetFileExtension(strFile)
    Case "ICO"
        Call SaveToIcon(strFile) 'アルファチェネルの設定あり
    Case "BMP"
        Call SaveToBmp(strFile) 'アルファチェネルの設定あり
    Case "PNG"
        Dim Gdip As New CGdiplus
        Call Gdip.CreateFromPixels(FPixels())
        Call Gdip.SaveToPNGFile(strFile)
    Case Else
        Call Err.Raise(513, , "ファイル形式エラー")
    End Select
End Sub

'*****************************************************************************
'[概要] BMP形式の画像をクリップボードから読込む
'[引数] なし
'[戻値] なし
'*****************************************************************************
Public Sub LoadBMPFromClipbord(Optional lngWidth As Long = 0, Optional lngHeight As Long = 0)
    If OpenClipboard(0) = 0 Then Exit Sub
On Error GoTo Finalization
    Dim hBitmap As LongPtr
    hBitmap = GetClipboardData(CF_BITMAP)
    If hBitmap = 0 Then GoTo Finalization
    
    Dim Gdip As New CGdiplus
    Call Gdip.CreateFromHBITMAP(hBitmap, lngWidth, lngHeight)
    Call Me.SetSize(Gdip.Width, Gdip.Height)
    Call Gdip.ToPixels(FPixels())
    Call SetNoAlpha
Finalization:
    Call CloseClipboard
    If Err.Number <> 0 Then Call Err.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
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

    Dim Stream As IUnknown
    If CreateStreamOnHGlobal(ByVal hGlobal, 0, Stream) <> S_OK Then
        Call Err.Raise(513, , "CreateStreamOnHGlobalエラー")
    End If
    
    Dim Gdip As New CGdiplus
    Call Gdip.CreateFromStream(Stream)
    Call Me.SetSize(Gdip.Width, Gdip.Height)
    Call Gdip.ToPixels(FPixels())
    Call SetNoAlpha
Finalization:
    Call CloseClipboard
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
    If Err.Number <> 0 Then Call Err.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'*****************************************************************************
'[概要] PNG形式の画像をクリップボードに書込む
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub SavePNGToClipbord()
    Dim Stream As IUnknown
    If CreateStreamOnHGlobal(0, 0, Stream) <> S_OK Then
        Call Err.Raise(513, , "CreateStreamOnHGlobalエラー")
    End If
    
    Dim Gdip As New CGdiplus
    Call Gdip.CreateFromPixels(FPixels())
    Call Gdip.ToPNGStream(Stream)
    
    Dim hGlobal As LongPtr
    Call GetHGlobalFromStream(Stream, hGlobal)
    If hGlobal = 0 Then
        Call Err.Raise(513, , "GetHGlobalFromStreamエラー")
    End If
    
    Dim CF_PNG As Long
    CF_PNG = RegisterClipboardFormat("PNG")
    If SetClipboardData(CF_PNG, hGlobal) = 0 Then
        Call Err.Raise(513, , "SetClipboardDataエラー")
    End If
End Sub

'*****************************************************************************
'[概要] BMP形式の画像をクリップボードに書込む
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub SaveBMPToClipbord()
    Dim Gdip As New CGdiplus
    Call Gdip.CreateFromPixels(FPixels())
    Call SetClipboardData(CF_BITMAP, Gdip.ToHBITMAP)
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
            FPixels(x, y) = CellToRGBQuad(objRange.Cells(y, x))
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
    
    '高速化のためいったん書式をクリア
    Call ClearRange(objRange.Resize(FHeight, FWidth))
    Dim x As Long, y As Long
    For y = 1 To FHeight
        For x = 1 To FWidth
            Call RGBQuadToCell(objRange.Cells(y, x), FPixels(x, y), False)
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
    If Err.Number <> 0 Then Call Err.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
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
            If RGBQuadToAlpha(FPixels(x, y + 1)) = 0 Then
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
    If Err.Number <> 0 Then Call Err.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
End Sub

'*****************************************************************************
'[概要] hBitmapからFPixelsにアルファチャネルありの色を設定する
'[引数] Bitmapハンドル
'[戻値] なし
'*****************************************************************************
Public Sub GetPixelsFromHBITMAP(ByVal hBitmap As LongPtr)
    Dim Gdip As New CGdiplus
    Call Gdip.CreateFromHBITMAP(hBitmap)
    Call Me.SetSize(Gdip.Width, Gdip.Height)
    Call Gdip.ToPixels(FPixels())
    Call SetNoAlpha
End Sub

'*****************************************************************************
'[概要] IPictureに設定する
'[引数] なし
'[戻値] IPicture
'*****************************************************************************
Public Function SetToIPicture() As IPicture
    Dim Gdip As New CGdiplus
    Call Gdip.CreateFromPixels(FPixels())
    Set SetToIPicture = Gdip.ToIPicture
End Function

'*****************************************************************************
'[概要] 塗りつぶす
'[引数] x,y:塗りつぶし開始位置、DstColor:塗りつぶし色
'[戻値] なし
'*****************************************************************************
Public Sub Fill(ByVal x As Long, ByVal y As Long, ByVal DstColor As Long)
    'ワーク用の2次元配列を確保(上下左右に壁用のマスを確保する)
    '格納する値(0:開始セルと同色のセル、1:塗りつぶし完了セル、9:対象外の色のセルまたは壁)
    ReDim Pixels(0 To FWidth + 1, 0 To FHeight + 1) As Byte
    Dim xx As Long, yy As Long
    '配列の初期化
    For xx = 0 To FWidth + 1
        For yy = 0 To FHeight + 1
            Pixels(xx, yy) = 9
        Next
    Next
    For xx = 1 To FWidth
        For yy = 1 To FHeight
            If FPixels(xx, yy) = FPixels(x, y) Then
                '開始セルと同じ色のセルに0を設定
                Pixels(xx, yy) = 0
            End If
        Next
    Next
    
'    Dim dblTime As Double
'    Dim cnt As Long
'    cnt = 0
'    Call MsgBox("開始")
'    dblTime = Timer()
    '塗りつぶし対象セル全てに1を設定する
    Call FillCell(x, y, Pixels())
'    Call MsgBox("実行時間：" & (Timer() - dblTime) * 1000 & " ミリ秒" & "　回数：" & cnt)
    
    '塗りつぶし対象セルに該当色を設定
    For xx = 1 To FWidth
        For yy = 1 To FHeight
            If Pixels(xx, yy) = 1 Then
                FPixels(xx, yy) = DstColor
            End If
        Next
    Next
End Sub

'*****************************************************************************
'[概要] 一筆書きを実行し、行詰れば次の起点となるセルを検索しそのセルから一筆書きを実行する
'       起点となるセルが見つからなくなるまで、一筆書きを続ける
'[引数] x,y:座標、Pixels():塗りつぶし用の配列
'[戻値] なし
'*****************************************************************************
Private Sub FillCell(ByVal x As Long, ByVal y As Long, ByRef Pixels() As Byte)
    '一筆書きを実行
    Call Onestroke(x, y, Pixels())
    
    Dim blnContinue As Boolean
    '起点となるセルが見つからなくなるまでLOOP
    Do While True
        blnContinue = False
        For y = 1 To FHeight
            For x = 1 To FWidth
                If Pixels(x, y) = 0 Then '開始セルと同じ色
                    '隣接するセルにすでに塗りつぶしたセルがあるか
                    If Pixels(x - 1, y) = 1 Or _
                       Pixels(x + 1, y) = 1 Or _
                       Pixels(x, y - 1) = 1 Or _
                       Pixels(x, y + 1) = 1 Then
                        '起点となるセルから一筆書きを実行
                        Call Onestroke(x, y, Pixels())
                        blnContinue = True
                    End If
                End If
            Next
        Next
        If Not blnContinue Then
            Exit Sub
        End If
    Loop
End Sub

'*****************************************************************************
'[概要] 一筆書きを行詰るまで実行する
'[引数] x,y:開始座標、Pixels():塗りつぶし用の配列
'[戻値] なし
'*****************************************************************************
Private Sub Onestroke(ByVal x As Long, ByVal y As Long, ByRef Pixels() As Byte)
    Do While True
        Pixels(x, y) = 1 '塗りつぶし

        '上に自分と同じ色があるか
        If Pixels(x, y - 1) = 0 Then
            y = y - 1
        '下に自分と同じ色があるか
        ElseIf Pixels(x, y + 1) = 0 Then
            y = y + 1
        '左に自分と同じ色があるか
        ElseIf Pixels(x - 1, y) = 0 Then
            x = x - 1
        '右に自分と同じ色があるか
        ElseIf Pixels(x + 1, y) = 0 Then
            x = x + 1
        Else
            '左上に自分と同じ色があるか　かつ　隣接セルに塗りつぶしが完了したセルがあるか
            If Pixels(x - 1, y - 1) = 0 And (Pixels(x - 1, y) = 1 Or Pixels(x, y - 1) = 1) Then
                x = x - 1: y = y - 1
            '左下に自分と同じ色があるか　かつ　隣接セルに塗りつぶしが完了したセルがあるか
            ElseIf Pixels(x - 1, y + 1) = 0 And (Pixels(x - 1, y) = 1 Or Pixels(x, y + 1) = 1) Then
                x = x - 1: y = y + 1
            '右上に自分と同じ色があるか　かつ　隣接セルに塗りつぶしが完了したセルがあるか
            ElseIf Pixels(x + 1, y - 1) = 0 And (Pixels(x + 1, y) = 1 Or Pixels(x, y - 1) = 1) Then
                x = x + 1: y = y - 1
            '右下に自分と同じ色があるか　かつ　隣接セルに塗りつぶしが完了したセルがあるか
            ElseIf Pixels(x + 1, y + 1) = 0 And (Pixels(x + 1, y) = 1 Or Pixels(x, y + 1) = 1) Then
                x = x + 1: y = y + 1
            Else
                '一筆書き終了
                Exit Sub
            End If
        End If
    Loop
End Sub

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
            FPixels(x, y) = FPixels(x, y) Or &HFF000000 '不透明
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
            If RGBQuadToAlpha(FPixels(x, y)) <> 0 Then
                HasAlpha = True
                Exit Function
            End If
        Next
    Next
End Function

