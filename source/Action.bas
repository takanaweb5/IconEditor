Attribute VB_Name = "Action"
Option Explicit

Private Declare PtrSafe Function GetKeyState Lib "user32" (ByVal lngVirtKey As Long) As Integer
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hDC As LongPtr) As Long

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As LongPtr) As LongPtr
Private Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal hDC As LongPtr) As Long
Private Declare PtrSafe Function GetDIBits Lib "gdi32" (ByVal aHDC As LongPtr, ByVal hBitmapptr As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, ByRef lpBits As Any, ByRef lpBI As Any, ByVal wUsage As Long) As Long
Private Const DIB_RGB_COLORS = 0&

Private Declare PtrSafe Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As LongPtr, ByVal nCount As Long, lpObject As Any) As Long
Private Type BITMAPINFOHEADER
    Size          As Long
    Width         As Long
    Height        As Long
    Planes        As Integer
    BitCount      As Integer
    Compression   As Long
    SizeImg     As Long
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
#If Win64 Then
    Bits As LongPtr
    Reserve As Long '構造体を64Bitの倍数にするため
#Else
    Bits As Long
#End If
End Type

Private Const LOGPIXELSX = 88
Private Const LOGPIXELSY = 90

Public FFormLoad As Boolean

'*****************************************************************************
'[概要] ImageMsoを取得
'[引数] なし
'[戻値] なし
'*****************************************************************************
Public Sub ImageMso取得()
    If CheckSelection <> E_Range Then
        Call MsgBox("画像を読込むセルを選択してください", vbCritical)
        Exit Sub
    End If
    
    Dim strImgMso As String
    strImgMso = InputBox("ImageMsoを入力してください" & vbCrLf & vbCrLf & "例 Copy")
    If strImgMso = "" Then Exit Sub
    
    'チェック
    On Error Resume Next
    Call CommandBars.GetImageMso(strImgMso, 32, 32)
    If Err.Number <> 0 Then
        Call MsgBox("ImageMsoが誤っています")
        Exit Sub
    End If
    
    Dim WidthAndHeight As Variant
    Dim strInput As String
    Dim lngWidth As Long
    Dim lngHeight As Long
    Do While True
        strInput = InputBox("幅,高さを入力してください" & vbCrLf & vbCrLf & "例 32,32")
        If strInput = "" Then
            lngWidth = 32
            lngHeight = 32
            Exit Do
        End If
        WidthAndHeight = Split(strInput, ",")
        If UBound(WidthAndHeight) = 1 Then
            If IsNumeric(WidthAndHeight(0)) And IsNumeric(WidthAndHeight(1)) Then
                lngWidth = WidthAndHeight(0)
                lngHeight = WidthAndHeight(1)
                Exit Do
            End If
        End If
    Loop
    
On Error GoTo ErrHandle
    Dim img As New CImage
    Call img.GetPixelsFromHBITMAP(CommandBars.GetImageMso(strImgMso, lngWidth, lngHeight).Handle)
    Application.ScreenUpdating = False
    Call img.SetPixelsToRange(ActiveCell)
    Call ActiveCell.Resize(img.Height, img.Width).Select
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[概要] 画像を読込む
'[引数] なし
'[戻値] なし
'*****************************************************************************
Public Sub 画像読込()
On Error GoTo ErrHandle
    If CheckSelection <> E_Range Then
        Call MsgBox("画像を読込むセルを選択してください")
        Exit Sub
    End If
    Dim vDBName As Variant
    vDBName = Application.GetOpenFilename("PNG,*.png,アイコン,*.ico,ビットマップ,*.bmp,全てのファイル,*.*")
    If vDBName = False Then
        Exit Sub
    End If
    
    Dim img As New CImage
    Call img.LoadFromFile(vDBName)
    If img.Width > 256 Or img.Height > 256 Then
        Call MsgBox("幅または高さが256Pixelを超えるファイルは読み込めません")
        Exit Sub
    End If
    Application.ScreenUpdating = False
    Call img.SetPixelsToRange(ActiveCell)
    Call ActiveCell.Resize(img.Height, img.Width).Select
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[概要] 画像を保存
'[引数] なし
'[戻値] なし
'*****************************************************************************
Public Sub 画像保存()
On Error GoTo ErrHandle
    If CheckSelection <> E_Range Then
        Call MsgBox("画像を選択してください")
        Exit Sub
    End If
    If Selection.Rows.Count = 1 And Selection.Columns.Count = 1 Then
        Call MsgBox("画像を選択してください")
        Exit Sub
    End If
    If Selection.Rows.Count > 256 Or Selection.Columns.Count > 256 Then
        Call MsgBox("幅または高さが256Pixelを超える画像は保存できません")
        Exit Sub
    End If
    
    Dim vDBName As Variant
    vDBName = Application.GetSaveAsFilename("", "PNG,*.png,アイコン,*.ico,ビットマップ,*.bmp,全てのファイル,*.*")
    If vDBName = False Then Exit Sub
    
    Dim img As New CImage
    Call img.GetPixelsFromRange(Selection)
    Call img.SaveImageToFile(vDBName)
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[概要] クリップボードの画像を保存
'[引数] なし
'[戻値] なし
'*****************************************************************************
Public Sub Clipbord画像保存()
On Error GoTo ErrHandle
    If Not ClipboardHasBitmap() Then
        Call MsgBox("クリップボードに画像がありません")
        Exit Sub
    End If
    
    Dim vDBName As Variant
    vDBName = Application.GetSaveAsFilename("", "PNG,*.png,アイコン,*.ico,ビットマップ,*.bmp,全てのファイル,*.*")
    If vDBName = False Then Exit Sub
    
    Dim img As New CImage
    Call img.LoadBMPFromClipbord
    Call img.SaveImageToFile(vDBName)
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[概要] 上下反転
'[引数] なし
'[戻値] なし
'*****************************************************************************
Public Sub 上下反転()
On Error GoTo ErrHandle
    If CheckSelection <> E_Range Then Exit Sub
    Dim img As New CImage
    Call img.GetPixelsFromRange(Selection)
    Call img.FlipHorizontal
    Application.ScreenUpdating = False
    Call img.SetPixelsToRange(Selection)
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[概要] 左右反転
'[引数] なし
'[戻値] なし
'*****************************************************************************
Public Sub 左右反転()
On Error GoTo ErrHandle
    If CheckSelection <> E_Range Then Exit Sub
    Dim img As New CImage
    Call img.GetPixelsFromRange(Selection)
    Call img.FlipVertical
    Application.ScreenUpdating = False
    Call img.SetPixelsToRange(Selection)
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[概要] 90度回転
'[引数] Mode 1:選択範囲を回転,2:クリップボードの領域を回転
'       Angle:90 or -90
'[戻値] なし
'*****************************************************************************
Public Sub 回転(ByVal lngMode As Long, ByVal lngAngle As Long)
On Error GoTo ErrHandle
    If CheckSelection <> E_Range Then Exit Sub
    
    Dim objCopyRange As Range
    If lngMode = 1 Then
        If Selection.Rows.Count <> Selection.Columns.Count Then
            Call MsgBox("幅と高さが不一致のため実行できません" & vbCrLf & "貼付コマンドの中の回転を実行してください")
            Exit Sub
        Else
            Set objCopyRange = Selection
        End If
    Else
        Set objCopyRange = GetCopyRange()
        If objCopyRange Is Nothing Then
            Call MsgBox("回転させる領域をコピーしてから実行してください")
            Exit Sub
        End If
    End If

    Dim img As New CImage
    Call img.GetPixelsFromRange(objCopyRange)
    Call img.Rotate(lngAngle)
    Call img.SetPixelsToRange(Selection)
    Call ActiveCell.Resize(img.Height, img.Width).Select
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[概要] クリップボードの画像を読込む
'[引数] なし
'[戻値] なし
'*****************************************************************************
Public Sub Clipbord画像読込()
On Error GoTo ErrHandle
    If Not ClipboardHasBitmap() Then
        Call MsgBox("クリップボードに画像がありません")
        Exit Sub
    End If
    If CheckSelection <> E_Range Then
        Call MsgBox("画像を読込むセルを選択してください")
        Exit Sub
    End If
    
    Dim img As New CImage
    Call img.LoadBMPFromClipbord
    Dim strDefault As String
    strDefault = img.Width & "," & img.Height
    
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim WidthAndHeight As Variant
    Dim strInput As String
    Dim objSelection
    Set objSelection = Selection
    Do While True
        strInput = InputBox("幅,高さを入力してください", , strDefault)
        If strInput = "" Then
            lngWidth = img.Width
            lngHeight = img.Height
            Exit Do
        End If
        WidthAndHeight = Split(strInput, ",")
        If UBound(WidthAndHeight) = 1 Then
            If IsNumeric(WidthAndHeight(0)) And IsNumeric(WidthAndHeight(1)) Then
                lngWidth = WidthAndHeight(0)
                lngHeight = WidthAndHeight(1)
                Exit Do
            End If
        End If
    Loop
    
    Call img.Resize(lngWidth, lngHeight)
    Application.ScreenUpdating = False
    Call img.SetPixelsToRange(Selection)
    Call ActiveCell.Resize(img.Height, img.Width).Select
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[概要] 画像に変換
'[引数] なし
'[戻値] なし
'*****************************************************************************
Public Sub 画像に変換()
On Error GoTo ErrHandle
    Dim img As New CImage
    If CheckSelection <> E_Range Then Exit Sub
    Call img.GetPixelsFromRange(Selection)
    Call img.SaveImageToClipbord
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[概要] Shapeを画像に変換して読込む
'[引数] なし
'[戻値] なし
'*****************************************************************************
Public Sub Shape読込()
On Error GoTo ErrHandle
    If CheckSelection <> E_Shape Then
        Call MsgBox("オートシェイプが選択されていません")
        Exit Sub
    End If
    Dim objSelection
    Set objSelection = Selection
    
    Dim objCell As Range
    Call ActiveCell.Select
    Set objCell = SelectCell("画像を読込むセルを選択してください", ActiveCell)
    If objCell Is Nothing Then
        Exit Sub
    End If
    
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim WidthAndHeight As Variant
    Dim strInput As String
    Do While True
        strInput = InputBox("幅,高さを64Pixel未満で入力してください")
        If strInput = "" Then
            Exit Do
        End If
        WidthAndHeight = Split(strInput, ",")
        If UBound(WidthAndHeight) = 1 Then
            If IsNumeric(WidthAndHeight(0)) And IsNumeric(WidthAndHeight(1)) Then
                lngWidth = WidthAndHeight(0)
                lngHeight = WidthAndHeight(1)
                If lngWidth <= 64 And lngHeight <= 64 Then
                    Exit Do
                End If
            End If
        End If
    Loop
    Dim objWkShape As Shape
    Set objWkShape = GroupShape(objSelection.ShapeRange(1))
    
    '72(ExcelのデフォルトのDPI),96(Windows画像のデフォルトのDPI)
    objWkShape.Width = (lngWidth - 1) * 72 / 96
    objWkShape.Height = (lngHeight - 1) * 72 / 96
    Call objWkShape.Copy
    
    Dim img As New CImage
    Call img.LoadBMPFromClipbord
    Call img.Resize(lngWidth, lngHeight)
    Application.ScreenUpdating = False
    Call img.SetPixelsToRange(objCell)
    Call objCell.Resize(img.Height, img.Width).Select
    If Not (objWkShape Is Nothing) Then
        Call objWkShape.Delete
    End If
Exit Sub
ErrHandle:
    If Not (objWkShape Is Nothing) Then
        Call objWkShape.Delete
    End If
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[概要] 回転しているShapeは幅と高さが変になるのでGroup化してまともになるようにする
'[引数] グループ化前のShape
'[戻値] グループ化後のShape
'*****************************************************************************
Private Function GroupShape(ByRef objShape As Shape) As Shape
    ReDim lngIDArray(1 To 2) As Variant
    'クローンを２つ作成しグループ化する
    With objShape.Duplicate
        .Top = objShape.Top
        .Left = objShape.Left
        lngIDArray(1) = .ID
    End With
    With objShape.Duplicate
        .Top = objShape.Top
        .Left = objShape.Left
        '透明にする
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        lngIDArray(2) = .ID
    End With
    Set GroupShape = GetShapeRangeFromID(lngIDArray).Group
End Function

'*****************************************************************************
'[ 関数名 ]　GetShapeRangeFromID
'[ 概  要 ]　ShpesオブジェクトのIDからShapeRangeオブジェクトを取得
'[ 引  数 ]　IDの配列
'[ 戻り値 ]　ShapeRangeオブジェクト
'*****************************************************************************
Private Function GetShapeRangeFromID(ByRef lngID As Variant) As ShapeRange
    Dim i As Long
    Dim j As Long
    Dim lngShapeID As Long
    ReDim lngArray(LBound(lngID) To UBound(lngID)) As Variant
    For j = 1 To ActiveSheet.Shapes.Count
        lngShapeID = ActiveSheet.Shapes(j).ID
        For i = LBound(lngID) To UBound(lngID)
            If lngShapeID = lngID(i) Then
                lngArray(i) = j
                Exit For
            End If
        Next
    Next
    Set GetShapeRangeFromID = ActiveSheet.Shapes.Range(lngArray)
End Function

'*****************************************************************************
'[概要] 透明色の強調
'[引数] なし
'[戻値] なし
'*****************************************************************************
Public Sub 透明色強調()
On Error GoTo ErrHandle
    If CheckSelection <> E_Range Then Exit Sub
    
    Dim objSelection As Range
    Dim objCell As Range
    Set objSelection = Selection
    If objSelection.Rows.Count > 256 Or objSelection.Columns.Count > 256 Then
        Call MsgBox("幅または高さが256マスを超える時は実行出来ません")
        Exit Sub
    End If
    
    '選択範囲の重複を排除
    Set objSelection = ReSelectRange(objSelection)
    
    Dim img As New CImage
    Dim objArea As Range
    For Each objArea In objSelection.Areas
        Call img.GetPixelsFromRange(objArea)
        Call img.SetPixelsToRange(objArea)
    Next
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[概要] 色の置換
'[引数] なし
'[戻値] なし
'*****************************************************************************
Public Sub 色の置換()
On Error GoTo ErrHandle
    Dim objSelection As Range
    If CheckSelection = E_Range Then
        Set objSelection = Selection
    Else
        Set objSelection = ActiveCell
    End If
    Dim objCanvas As Range
    Set objCanvas = SelectCell("キャンバスの範囲を選択してください", objSelection)
    If objCanvas Is Nothing Then
        Exit Sub
    Else
        If objCanvas.Rows.Count > 256 Or objCanvas.Columns.Count > 256 Then
            Call MsgBox("幅または高さが256マスを超える時は実行出来ません")
            Exit Sub
        End If
    End If
    
    Dim objSrcCell As Range
    Set objSrcCell = SelectCell("置換前の色のセルを選択してください", ActiveCell)
    If objSrcCell Is Nothing Then
        Exit Sub
    End If
    Dim objDstCell As Range
    Set objDstCell = SelectCell("置換後の色のセルを選択してください", ActiveCell)
    If objDstCell Is Nothing Then
        Exit Sub
    End If
    
    '選択範囲の重複を排除
    Set objCanvas = ReSelectRange(objCanvas)
    
    Dim img As New CImage
    Dim lngColor As Long
    lngColor = img
    Dim objArea As Range
    For Each objArea In objCanvas
        If img.SameCellColor(objSrcCell(1), objArea) Then
            Debug.Print objArea.Address(0, 0)
            Call img.ChangeCellColor(objDstCell(1), objArea)
        End If
    Next
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[概要] 同じ（または相違する）色のセルを選択
'[引数] True:同じ色のセルを選択、False:違う色のセルを選択
'[戻値] なし
'*****************************************************************************
Public Sub 同色選択(ByVal blnSameColor As Boolean)
On Error GoTo ErrHandle
    Dim objSelection As Range
    If CheckSelection = E_Range Then
        Set objSelection = Selection
    Else
        Set objSelection = ActiveCell
    End If
    Dim objCanvas As Range
    Set objCanvas = SelectCell("キャンバスの範囲を選択してください", objSelection)
    If objCanvas Is Nothing Then
        Exit Sub
    Else
        If objCanvas.Rows.Count > 256 Or objCanvas.Columns.Count > 256 Then
            Call MsgBox("幅または高さが256マスを超える時は実行出来ません")
            Exit Sub
        End If
    End If
    
    Dim objColorCell As Range
    Dim strMsg As String
    If blnSameColor Then
        strMsg = "選択したい色のセルを選択してください"
    Else
        strMsg = "選択したくない色のセルを選択してください"
    End If
    Set objColorCell = SelectCell(strMsg, ActiveCell)
    If objColorCell Is Nothing Then
        Exit Sub
    End If
    
    Dim img As New CImage
    Dim objNewSelection As Range
    Dim objCell As Range
    For Each objCell In objCanvas
        If img.SameCellColor(objColorCell, objCell) = blnSameColor Then
            If objNewSelection Is Nothing Then
                Set objNewSelection = objCell
            Else
                Set objNewSelection = Application.Union(objNewSelection, objCell)
            End If
        End If
    Next
    If objNewSelection Is Nothing Then
        Call MsgBox("該当のセルはありませんでした")
    Else
        Call objNewSelection.Select
    End If
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[概要] 選択セルの反転など
'[引数] なし
'[戻値] なし
'*****************************************************************************
Public Sub 選択反転等()
    Static strLastSheet   As String '前回の領域の復元用
    Static strLastAddress As String '前回の領域の復元用
On Error GoTo ErrHandle
    Dim objSelection As Range
    Dim objUnSelect  As Range
    Dim objRange As Range
    Dim enmUnselectMode As EUnselectMode
    If CheckSelection() <> E_Range Then
        Exit Sub
    End If
    Set objSelection = Selection
    
    '取消領域を選択させる
    With frmUnSelect
        '前回の復元用
        Call .SetLastSelect(strLastSheet, strLastAddress)
        'フォームを表示
        Call .Show
        'キャンセル時
        If FFormLoad = False Then
            Exit Sub
        End If
        enmUnselectMode = .Mode
        Select Case (enmUnselectMode)
        Case E_Unselect, E_Reverse, E_Intersect, E_Union
            Set objUnSelect = .SelectRange
        End Select
        Call Unload(frmUnSelect)
    End With

    Select Case (enmUnselectMode)
    Case E_Unselect  '取消し
        Set objRange = MinusRange(objSelection, objUnSelect)
    Case E_Reverse   '反転
        Set objRange = UnionRange(MinusRange(objSelection, objUnSelect), MinusRange(objUnSelect, objSelection))
    Case E_Intersect '絞り込み
        Set objRange = IntersectRange(objSelection, objUnSelect)
    Case E_Union     '追加
        Set objRange = UnionRange(objSelection, objUnSelect)
    End Select
    
    If Not (objRange Is Nothing) Then
        Call objRange.Select
    End If
ErrHandle:
    strLastSheet = ActiveSheet.Name
    strLastAddress = Selection.Address(False, False)
    If FFormLoad Then
        Call Unload(frmUnSelect)
    End If
End Sub

'*****************************************************************************
'[概要] 色の貼付け
'[引数] なし
'[戻値] なし
'*****************************************************************************
Public Sub 貼付け()
On Error GoTo ErrHandle
    If CheckSelection <> E_Range Then Exit Sub
    Dim objCopyRange  As Range
    Set objCopyRange = GetCopyRange()
    If objCopyRange Is Nothing Then Exit Sub
    If objCopyRange.Rows.Count > 256 Or objCopyRange.Columns.Count > 256 Then
        Call MsgBox("幅または高さが256マスを超える時は実行出来ません")
        Exit Sub
    End If
    
    Dim objSelection As Range
    Set objSelection = Selection
    Dim objColorCell As Range
    Dim ColorFlg As Long
    Dim lngMode As Long
    If objCopyRange.Count > 1 Then
        If objSelection.Areas.Count > 1 Then
            Call MsgBox("このコマンドは複数の選択範囲に対して実行できません")
            Exit Sub
        End If
        If FChecked(1) Then
            Set objColorCell = SelectCell("対象色のセルを選択してください", ActiveCell)
            If objColorCell Is Nothing Then Exit Sub
            lngMode = 1
        End If
        If FChecked(2) Then
            Set objColorCell = SelectCell("除外対象の色のセルを選択してください", ActiveCell)
            If objColorCell Is Nothing Then Exit Sub
            lngMode = 2
        End If
    End If
    
    Dim img As New CImage
    Application.ScreenUpdating = False
    Call img.GetPixelsFromRange(objCopyRange)
    If objCopyRange.Count > 1 Then
        Call img.SetPixelsToRange(objSelection, lngMode, objColorCell, FChecked(3))
    Else
        Dim objCell As Range
        For Each objCell In objSelection
            Call img.SetPixelsToRange(objCell)
        Next
    End If
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[概要] クリップボードに画像を設定する
'[引数] なし
'[戻値] なし
'*****************************************************************************
Public Sub Clipbord画像設定()
On Error GoTo ErrHandle
    Dim img As New CImage
    Call img.GetPixelsFromRange(Selection)
    Call img.SaveImageToClipbord
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub


'*****************************************************************************
'[概要] 色のARGBを増減させる
'[引数] 1:増加、-1:減少
'[戻値] なし
'*****************************************************************************
Public Sub 色増減(ByVal lngUp As Long)
On Error GoTo ErrHandle
    If CheckSelection <> E_Range Then Exit Sub
    If FChecked(3) Or FChecked(4) Or FChecked(5) Or FChecked(6) Or FChecked(7) Then
    Else
        Call MsgBox("RGBおよびアルファ値のいずれもチェックされていません")
        Exit Sub
    End If
    If Selection.Rows.Count > 256 Or Selection.Columns.Count > 256 Then
        Call MsgBox("幅または高さが256マスを超える時は実行出来ません")
        Exit Sub
    End If

    If GetKeyState(vbKeyControl) < 0 Then
        lngUp = lngUp * 1
    Else
        lngUp = lngUp * 16
    End If
    
    '選択範囲の重複を排除
    Dim objSelection As Range
    Set objSelection = ReSelectRange(Selection)
    
    Dim img As New CImage
    Dim objArea As Range
    For Each objArea In objSelection.Areas
        Call img.GetPixelsFromRange(objArea)
        Call img.AdjustColor(lngUp, FChecked(4), FChecked(5), FChecked(6), FChecked(7))
        Call img.SetPixelsToRange(objArea)
    Next
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[概要] 画像をIPictureに変換する
'[引数] なし
'[戻値] IPicture
'*****************************************************************************
Public Function Getサンプル画像() As IPicture
On Error GoTo ErrHandle
    Dim img As New CImage
    Call img.GetPixelsFromRange(Selection)
    Set Getサンプル画像 = img.SetToIPicture
Exit Function
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Function

'*****************************************************************************
'[概要] 塗潰し
'[引数] なし
'[戻値] なし
'*****************************************************************************
Public Sub 塗潰し()
On Error GoTo ErrHandle
    Dim objSelection As Range
    If CheckSelection = E_Range Then
        Set objSelection = Selection.Areas(1)
    Else
        Set objSelection = ActiveCell
    End If
    Dim objCanvas As Range
    Set objCanvas = SelectCell("キャンバスの範囲を選択してください", objSelection)
    If objCanvas Is Nothing Then
        Exit Sub
    Else
        If objCanvas.Rows.Count > 256 Or objCanvas.Columns.Count > 256 Then
            Call MsgBox("幅または高さが256マスを超える時は実行出来ません")
            Exit Sub
        End If
    End If
    
    Dim objColorCell As Range
    Set objColorCell = SelectCell("塗潰す色のセルを選択してください", ActiveCell)
    If objColorCell Is Nothing Then
        Exit Sub
    End If
    
    Dim objStartCell As Range
    Set objStartCell = SelectCell("塗潰し開始セルを選択してください", ActiveCell)
    If objStartCell Is Nothing Then
        Exit Sub
    End If
    
    If Intersect(objCanvas, objStartCell) Is Nothing Then
        Call MsgBox("塗潰し開始セルが、キャンバスの内側にありません")
        Exit Sub
    End If
    
    Dim img As New CImage
    Call img.GetPixelsFromRange(objCanvas)
    Call img.Fill(objStartCell.Column - objCanvas.Column + 1, _
                  objStartCell.Row - objCanvas.Row + 1, _
                  objColorCell)
    Call img.SetPixelsToRange(objCanvas)
    
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[概要] 領域と領域の重なる領域を取得する
'[引数] 対象領域(Nothingも可)
'[戻値] objRange1 ∩ objRange2
'*****************************************************************************
Private Function IntersectRange(ByRef objRange1 As Range, ByRef objRange2 As Range) As Range
    Select Case True
    Case (objRange1 Is Nothing) Or (objRange2 Is Nothing)
        Set IntersectRange = Nothing
    Case Else
        Set IntersectRange = Intersect(objRange1, objRange2)
    End Select
End Function

'*****************************************************************************
'[概要] 領域に領域を加える
'[引数] 対象領域(Nothingも可)
'[戻値] objRange1 ∪ objRange2
'*****************************************************************************
Private Function UnionRange(ByRef objRange1 As Range, ByRef objRange2 As Range) As Range
    Select Case True
    Case (objRange1 Is Nothing) And (objRange2 Is Nothing)
        Set UnionRange = Nothing
    Case (objRange1 Is Nothing)
        Set UnionRange = objRange2
    Case (objRange2 Is Nothing)
        Set UnionRange = objRange1
    Case Else
        Set UnionRange = Union(objRange1, objRange2)
    End Select
End Function

'*****************************************************************************
'[概要] 領域から領域を、除外する
'       Ａ－Ｂ = Ａ∩!Ｂ
'       !Ｂ = !(B1∪B2∪B3...∪Bn) = !B1∩!B2∩!B3...∩!Bn
'[引数] 対象領域
'[戻値] objRange1 － objRange2
'*****************************************************************************
Private Function MinusRange(ByRef objRange1 As Range, ByRef objRange2 As Range) As Range
    Dim objRounds As Range
    Dim i As Long
    
    If objRange2 Is Nothing Then
        Set MinusRange = objRange1
        Exit Function
    End If
    
    '除外する領域の数だけループ
    '!Ｂ = !B1∩!B2∩!B3.....∩!Bn
    Set objRounds = ReverseRange(objRange2.Areas(1))
    For i = 2 To objRange2.Areas.Count
        Set objRounds = IntersectRange(objRounds, ReverseRange(objRange2.Areas(i)))
    Next
    
    'Ａ∩!Ｂ
    Set MinusRange = IntersectRange(objRange1, objRounds)
End Function

'*****************************************************************************
'[概要] 領域を反転する
'[引数] 対象領域
'[戻値] !objRange
'*****************************************************************************
Private Function ReverseRange(ByRef objRange As Range) As Range
    Dim i As Long
    Dim objRound(1 To 4) As Range
    
    With objRange.Parent
        On Error Resume Next
        '選択領域より上の領域すべて
        Set objRound(1) = .Range(.Rows(1), _
                                 .Rows(objRange.Row - 1))
        '選択領域より下の領域すべて
        Set objRound(2) = .Range(.Rows(objRange.Row + objRange.Rows.Count), _
                                 .Rows(Rows.Count))
        '選択領域より左の領域すべて
        Set objRound(3) = .Range(.Columns(1), _
                                 .Columns(objRange.Column - 1))
        '選択領域より右の領域すべて
        Set objRound(4) = .Range(.Columns(objRange.Column + objRange.Columns.Count), _
                                 .Columns(Columns.Count))
        On Error GoTo 0
    End With
    
    '選択領域以外の領域を設定
    For i = 1 To 4
        Set ReverseRange = UnionRange(ReverseRange, objRound(i))
    Next
End Function

'*****************************************************************************
'[概要] 領域の重複を省いた領域を取得
'[引数] 対象領域
'[戻値] 領域の重複を省いた領域
'*****************************************************************************
Private Function ReSelectRange(ByRef objRange As Range) As Range
    Dim objArrange(1 To 3) As Range
    With objRange
        On Error Resume Next
        Set objArrange(1) = .SpecialCells(xlCellTypeConstants)
        Set objArrange(2) = .SpecialCells(xlCellTypeFormulas)
        Set objArrange(3) = .SpecialCells(xlCellTypeBlanks)
        On Error GoTo 0
    End With

    Dim i As Long
    For i = 1 To 3
        Set ReSelectRange = UnionRange(ReSelectRange, objArrange(i))
    Next
End Function

