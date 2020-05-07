Attribute VB_Name = "Action"
Option Explicit

Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
Private Declare PtrSafe Function GetKeyState Lib "user32" (ByVal lngVirtKey As Long) As Integer
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
        strInput = InputBox("幅,高さを入力してください" & vbCrLf & vbCrLf & "例 32,32", , "32,32")
        If strInput = "" Then
            Exit Sub
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
    Call SaveUndoInfo(ActiveCell.Resize(img.Height, img.Width))
    Call img.SetPixelsToRange(ActiveCell)
    Call ActiveCell.Resize(img.Height, img.Width).Select
    Call SetOnUndo("ImageMso取得")
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
    Call img.LoadImageFromFile(vDBName)
    Application.ScreenUpdating = False
    Call SaveUndoInfo(ActiveCell.Resize(img.Height, img.Width))
    Call img.SetPixelsToRange(ActiveCell)
    Call ActiveCell.Resize(img.Height, img.Width).Select
    Call SetOnUndo("画像読込")
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
    If Selection.Rows.Count = Rows.Count Or Selection.Columns.Count = Columns.Count Then
        Call MsgBox("すべての行または列の選択時は実行出来ません")
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
    If Selection.Rows.Count = Rows.Count Or Selection.Columns.Count = Columns.Count Then
        Call MsgBox("すべての行または列の選択時は実行出来ません")
        Exit Sub
    End If
    
    Dim img As New CImage
    Call img.GetPixelsFromRange(Selection)
    Call img.FlipHorizontal
    Application.ScreenUpdating = False
    Call SaveUndoInfo(Selection)
    Call img.SetPixelsToRange(Selection)
    Call SetOnUndo("反転")
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
    If Selection.Rows.Count = Rows.Count Or Selection.Columns.Count = Columns.Count Then
        Call MsgBox("すべての行または列の選択時は実行出来ません")
        Exit Sub
    End If
    
    Dim img As New CImage
    Call img.GetPixelsFromRange(Selection)
    Call img.FlipVertical
    Application.ScreenUpdating = False
    Call SaveUndoInfo(Selection)
    Call img.SetPixelsToRange(Selection)
    Call SetOnUndo("反転")
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
    If Selection.Rows.Count = Rows.Count Or Selection.Columns.Count = Columns.Count Then
        Call MsgBox("すべての行または列の選択時は実行出来ません")
        Exit Sub
    End If
    
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
    Application.ScreenUpdating = False
    Call SaveUndoInfo(ActiveCell.Resize(img.Height, img.Width))
    Call img.SetPixelsToRange(Selection)
    Call ActiveCell.Resize(img.Height, img.Width).Select
    Call SetOnUndo("回転")
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
            Exit Sub
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
    Call SaveUndoInfo(ActiveCell.Resize(img.Height, img.Width))
    Call img.SetPixelsToRange(Selection)
    Call ActiveCell.Resize(img.Height, img.Width).Select
    Call SetOnUndo("画像読込")
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
    If CheckSelection <> E_Range Then Exit Sub
    If Selection.Rows.Count = Rows.Count Or Selection.Columns.Count = Columns.Count Then
        Call MsgBox("すべての行または列の選択時は実行出来ません")
        Exit Sub
    End If
    
    Dim img As New CImage
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
        strInput = InputBox("幅,高さを" & MAX_WIDTH & "Pixel未満で入力してください")
        If strInput = "" Then
            Exit Do
        End If
        WidthAndHeight = Split(strInput, ",")
        If UBound(WidthAndHeight) = 1 Then
            If IsNumeric(WidthAndHeight(0)) And IsNumeric(WidthAndHeight(1)) Then
                lngWidth = WidthAndHeight(0)
                lngHeight = WidthAndHeight(1)
                If lngWidth <= MAX_WIDTH And lngHeight <= MAX_HEIGHT Then
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
    Call SaveUndoInfo(objCell.Resize(img.Height, img.Width))
    Call img.SetPixelsToRange(objCell)
    Call objCell.Resize(img.Height, img.Width).Select
    If Not (objWkShape Is Nothing) Then
        Call objWkShape.Delete
    End If
    Call SetOnUndo("画像読込")
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
    If Selection.Rows.Count = Rows.Count Or Selection.Columns.Count = Columns.Count Then
        Call MsgBox("すべての行または列の選択時は実行出来ません")
        Exit Sub
    End If
    
    '選択範囲の重複を排除
    Dim objSelection As Range
    Set objSelection = ReSelectRange(Selection)
    
    Application.ScreenUpdating = False
    Call SaveUndoInfo(Selection)
    Dim objCell As Range
    For Each objCell In objSelection
        Call ColorToCell(objCell, CellToColor(objCell))
    Next
    Call SetOnUndo("透明色強調")
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
        If objCanvas.Rows.Count = Rows.Count Or objCanvas.Columns.Count = Columns.Count Then
            Call MsgBox("すべての行または列の選択時は実行出来ません")
            Exit Sub
        End If
    End If
    '選択範囲の重複を排除
    Set objCanvas = ReSelectRange(objCanvas)
    
    Dim objCell As Range
    Dim SrcColor As TRGBQuad
    Set objCell = SelectCell("置換前の色のセルを選択してください", ActiveCell)
    If objCell Is Nothing Then
        Exit Sub
    End If
    SrcColor = CellToColor(objCell(1))
    
    Dim DstColor As TRGBQuad
    Set objCell = SelectCell("置換後の色のセルを選択してください", ActiveCell)
    If objCell Is Nothing Then
        Exit Sub
    End If
    DstColor = CellToColor(objCell(1))
    
    Application.ScreenUpdating = False
    Call SaveUndoInfo(Selection)
    For Each objCell In objCanvas
        If SameColor(SrcColor, CellToColor(objCell)) Then
            Call ColorToCell(objCell, DstColor)
        End If
    Next
    Call Selection.Select
    Call SetOnUndo("色の置換")
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
    Dim objCell As Range
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
        If objCanvas.Rows.Count = Rows.Count Or objCanvas.Columns.Count = Columns.Count Then
            Call MsgBox("すべての行または列の選択時は実行出来ません")
            Exit Sub
        End If
    End If
    
    Dim SelectColor  As TRGBQuad
    Dim strMsg As String
    If blnSameColor Then
        strMsg = "選択したい色のセルを選択してください"
    Else
        strMsg = "選択したくない色のセルを選択してください"
    End If
    Set objCell = SelectCell(strMsg, ActiveCell)
    If objCell Is Nothing Then
        Exit Sub
    End If
    SelectColor = CellToColor(objCell(1))
    
    Dim objNewSelection As Range
    For Each objCell In objCanvas
        If SameColor(SelectColor, CellToColor(objCell)) = blnSameColor Then
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
    If Selection.Rows.Count = Rows.Count Or Selection.Columns.Count = Columns.Count Then
        Call MsgBox("すべての行または列の選択時は実行出来ません")
        Exit Sub
    End If
    Dim objCopyRange  As Range
    Set objCopyRange = GetCopyRange()
    If objCopyRange Is Nothing Then Exit Sub
    If objCopyRange.Rows.Count = Rows.Count Or objCopyRange.Columns.Count = Columns.Count Then
        Call MsgBox("すべての行または列の選択時は実行出来ません")
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
    
    Application.ScreenUpdating = False
    If objCopyRange.Count = 1 Then
        Dim DstColor As TRGBQuad
        DstColor = CellToColor(objCopyRange)
        Call SaveUndoInfo(objSelection)
        Dim objCell As Range
        For Each objCell In objSelection
            Call ColorToCell(objCell, DstColor)
        Next
    Else
        Dim img As New CImage
        Call img.GetPixelsFromRange(objCopyRange)
        Call SaveUndoInfo(objSelection.Resize(img.Height, img.Width))
        Call img.SetPixelsToRange(objSelection, lngMode, objColorCell, FChecked(3))
        Call objSelection.Resize(img.Height, img.Width).Select
    End If
    Call SetOnUndo("貼付け")
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
    If Selection.Rows.Count = Rows.Count Or Selection.Columns.Count = Columns.Count Then
        Call MsgBox("すべての行または列の選択時は実行出来ません")
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
    
    Dim objCell As Range
    Dim ARGB As TRGBQuad
    Application.ScreenUpdating = False
    Call SaveUndoInfo(Selection, "色調整")
    For Each objCell In objSelection
        ARGB = AdjustColor(CellToColor(objCell), lngUp, FChecked(4), FChecked(5), FChecked(6), FChecked(7))
        Call ColorToCell(objCell, ARGB)
    Next
    Call Selection.Select
    Call SetOnUndo("色調整")
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
        If objCanvas.Rows.Count = Rows.Count Or objCanvas.Columns.Count = Columns.Count Then
            Call MsgBox("すべての行または列の選択時は実行出来ません")
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
    Application.ScreenUpdating = False
    Call SaveUndoInfo(objCanvas)
    Call img.SetPixelsToRange(objCanvas)
    Call SetOnUndo("塗潰し")
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

