Attribute VB_Name = "Action"
Option Explicit
Option Private Module

Public Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
Private Declare PtrSafe Function GetKeyState Lib "user32" (ByVal lngVirtKey As Long) As Integer
Public FFormLoad As Boolean

'*****************************************************************************
'[概要] クリア
'[引数] なし
'[戻値] なし
'*****************************************************************************
Public Sub クリア()
On Error GoTo ErrHandle
    If CheckSelection <> E_Range Then Exit Sub
    Application.ScreenUpdating = False
    Call SaveUndoInfo(Selection)
    Call ClearRange(Selection)
'    Call ColorToCell(Selection, OleColorToARGB(0, 0), True)
    Call SetOnUndo("クリア")
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

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
    Dim objSelection As Range
    Set objSelection = Selection
    
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
    
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim DestRange As Range
    If 1 < objSelection.Columns.Count And objSelection.Columns.Count <= 64 And _
       1 < objSelection.Rows.Count And objSelection.Rows.Count <= 64 Then
        lngWidth = objSelection.Columns.Count
        lngHeight = objSelection.Rows.Count
        Set DestRange = objSelection
    Else
        Dim WidthAndHeight As Variant
        Dim strInput As String
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
        Set DestRange = ActiveCell.Resize(lngHeight, lngWidth)
    End If
    If 16 <= lngWidth And lngWidth <= 64 And _
       16 <= lngHeight And lngHeight <= 64 Then
    Else
        Call MsgBox("幅および高さは 16～64 で指定してください")
        Exit Sub
    End If

On Error GoTo ErrHandle
    Dim img As New CImage
    Call img.GetPixelsFromHBITMAP(CommandBars.GetImageMso(strImgMso, lngWidth, lngHeight).Handle)
    Application.ScreenUpdating = False
    Call SaveUndoInfo(DestRange)
    Call img.SetPixelsToRange(DestRange)
    Call DestRange.Select
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
    If Selection.Rows.Count = Rows.Count Or Selection.Columns.Count = Columns.Count Then
        Call MsgBox("すべての行または列の選択時は実行出来ません")
        Exit Sub
    End If
    Dim objSelection As Range
    Set objSelection = Selection
    
    Dim vDBName As Variant
    vDBName = Application.GetOpenFilename("PNG,*.png,アイコン,*.ico,ビットマップ,*.bmp,全てのファイル,*.*")
    If vDBName = False Then
        Exit Sub
    End If
    
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim DestRange As Range
    If 1 < objSelection.Columns.Count And _
       1 < objSelection.Rows.Count Then
        lngWidth = objSelection.Columns.Count
        lngHeight = objSelection.Rows.Count
        Set DestRange = objSelection
    End If
    
    Dim img As New CImage
    Call img.LoadImageFromFile(vDBName)
    If DestRange Is Nothing Then
        Set DestRange = ActiveCell.Resize(img.Height, img.Width)
    Else
        If lngWidth > img.Width Or lngHeight > img.Height Then
            Set DestRange = DestRange.Resize(img.Height, img.Width)
        Else
            Call img.Resize(lngWidth, lngHeight)
        End If
    End If
    
    Application.ScreenUpdating = False
    Call SaveUndoInfo(DestRange)
    Call img.SetPixelsToRange(DestRange)
    Call DestRange.Select
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
    If Selection.Rows.Count = Rows.Count Or Selection.Columns.Count = Columns.Count Then
        Call MsgBox("すべての行または列の選択時は実行出来ません")
        Exit Sub
    End If
    Dim objSelection As Range
    Set objSelection = Selection
    
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim DestRange As Range
    If 1 < objSelection.Columns.Count And _
       1 < objSelection.Rows.Count Then
        lngWidth = objSelection.Columns.Count
        lngHeight = objSelection.Rows.Count
        Set DestRange = objSelection
    End If
    
    Dim img As New CImage
    Call img.LoadBMPFromClipbord
    If DestRange Is Nothing Then
        Set DestRange = ActiveCell.Resize(img.Height, img.Width)
    Else
        If lngWidth > img.Width Or lngHeight > img.Height Then
            Set DestRange = DestRange.Resize(img.Height, img.Width)
        Else
            Call img.Resize(lngWidth, lngHeight)
        End If
    End If
    
    Application.ScreenUpdating = False
    Call SaveUndoInfo(DestRange)
    Call img.SetPixelsToRange(DestRange)
    Call DestRange.Select
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
    
    Dim DestRange As Range
    Call ActiveCell.Select
    Set DestRange = SelectCell("画像を読込むセルを選択してください", ActiveCell)
    If DestRange Is Nothing Then
        Exit Sub
    End If
    
    Dim lngWidth As Long
    Dim lngHeight As Long
    If 1 < DestRange.Columns.Count And DestRange.Columns.Count <= 64 And _
       1 < DestRange.Rows.Count And DestRange.Rows.Count <= 64 Then
        lngWidth = DestRange.Columns.Count
        lngHeight = DestRange.Rows.Count
    Else
        Dim WidthAndHeight As Variant
        Dim strInput As String
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
        Set DestRange = DestRange.Resize(lngHeight, lngWidth)
    End If
    
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
    Call SaveUndoInfo(DestRange)
    Call img.SetPixelsToRange(DestRange)
    Call DestRange.Select
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
        Call ColorToCell(objCell, CellToColor(objCell), True)
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
    If CheckSelection <> E_Range Then
        Call MsgBox("対象の範囲を選択してから実行してください")
        Exit Sub
    End If
    Set objSelection = Selection
    If objSelection.Rows.Count = Rows.Count Or objSelection.Columns.Count = Columns.Count Then
        Call MsgBox("すべての行または列の選択時は実行出来ません")
        Exit Sub
    End If
    If objSelection.Count = 1 Then
        Call MsgBox("対象の範囲を選択してから実行してください")
        Exit Sub
    End If
    
    Dim objCell As Range
    Dim SrcColor As TRGBQuad
    Set objCell = SelectCell("置換前の色のセルを選択してください", ActiveCell)
    If objCell Is Nothing Then
        Exit Sub
    End If
    If Intersect(objSelection, objCell) Is Nothing Then
        Call MsgBox("セルが、対象範囲の内側にありません")
        Exit Sub
    End If
    SrcColor = CellToColor(objCell(1))
    
    Dim DstColor As TRGBQuad
    Set objCell = SelectCell("置換後の色のセルを選択してください", ActiveCell)
    If objCell Is Nothing Then
        Exit Sub
    End If
    DstColor = CellToColor(objCell(1))
    
    '選択範囲の重複を排除
    Dim objCanvas As Range
    Set objCanvas = ReSelectRange(objSelection)
    
    '置換対象セルを取得
    Dim objRange As Range
    For Each objCell In objCanvas
        If SameColor(SrcColor, CellToColor(objCell)) Then
            Set objRange = UnionRange(objRange, objCell)
        End If
    Next
    If objRange Is Nothing Then Exit Sub
    
    Application.ScreenUpdating = False
    Call SaveUndoInfo(Selection)
    Call ColorToCell(objRange, DstColor, True)
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
    If CheckSelection <> E_Range Then
        Call MsgBox("対象の範囲を選択してから実行してください")
        Exit Sub
    End If
    Set objSelection = Selection
    If objSelection.Rows.Count = Rows.Count Or objSelection.Columns.Count = Columns.Count Then
        Call MsgBox("すべての行または列の選択時は実行出来ません")
        Exit Sub
    End If
    If objSelection.Count = 1 Then
        Call MsgBox("対象の範囲を選択してから実行してください")
        Exit Sub
    End If
    
    Dim SelectColor  As TRGBQuad
    Dim strMsg As String
    If blnSameColor Then
        strMsg = "選択したい色のセルを選択してください"
    Else
        strMsg = "選択したくない色のセルを選択してください"
    End If
    Dim objCell As Range
    Set objCell = SelectCell(strMsg, ActiveCell)
    If objCell Is Nothing Then
        Exit Sub
    End If
'    If Intersect(objSelection, objCell) Is Nothing Then
'        Call MsgBox("セルが、対象範囲の内側にありません")
'        Exit Sub
'    End If
    SelectColor = CellToColor(objCell(1))
    
    '選択範囲の重複を排除
    Dim objCanvas As Range
    Set objCanvas = ReSelectRange(objSelection)
    
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
        Call ReSelectRange(objRange).Select
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
    Dim objCell As Range
    Dim Color   As TRGBQuad
    If objCopyRange.Count > 1 Then
        If objSelection.Areas.Count > 1 Then
            Call MsgBox("このコマンドは複数の選択範囲に対して実行できません")
            Exit Sub
        End If
        If GetTmpControl("C1").State Then
            Set objCell = SelectCell("対象色のセルを選択してください", ActiveCell)
            If objCell Is Nothing Then Exit Sub
            If Intersect(objCopyRange, objCell) Is Nothing Then
                Call MsgBox("セルが、対象範囲の内側にありません")
                Exit Sub
            End If
            Color = CellToColor(objCell)
        End If
        If GetTmpControl("C2").State Then
            Set objCell = SelectCell("除外対象の色のセルを選択してください", ActiveCell)
            If objCell Is Nothing Then Exit Sub
            If Intersect(objCopyRange, objCell) Is Nothing Then
                Call MsgBox("セルが、対象範囲の内側にありません")
                Exit Sub
            End If
            Color = CellToColor(objCell)
        End If
    End If
    
    '貼付け先の領域
    Dim objDestRange As Range
    If objCopyRange.Count = 1 Then
        Set objDestRange = objSelection
    Else
        Set objDestRange = objSelection.Resize(objCopyRange.Rows.Count, objCopyRange.Columns.Count)
    End If
    
    Application.ScreenUpdating = False
    Call SaveUndoInfo(objDestRange)
    If objCopyRange.Count = 1 Then
        Dim DstColor As TRGBQuad
        DstColor = CellToColor(objCopyRange)
        Call ColorToCell(objDestRange, DstColor, True)
    Else
        If GetTmpControl("C1").State Or GetTmpControl("C2").State Then
            Call PasteSub(Color, objCopyRange, objDestRange)
        Else
            Dim img As New CImage
            Call img.GetPixelsFromRange(objCopyRange)
            Call img.SetPixelsToRange(objDestRange)
        End If
    End If
    Call objDestRange.Select
    Call SetOnUndo("貼付け")
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub


'*****************************************************************************
'[概要] 特定の色だけ または 特定の色を除く貼付け
'[引数] 対象または対象外とする色,Copy元のRange,貼付け先のRange
'[戻値] なし
'*****************************************************************************
Private Sub PasteSub(ByRef Color As TRGBQuad, ByRef objCopyRange As Range, ByRef objDestRange As Range)
    Dim objSameRange  As Range  '同じ色のセル
    Dim objDiffRange  As Range  '違う色のセル
    Dim objCell As Range
    
    '選択色と同じ色のセルと違う色のセルを取得
    Dim i As Long
    For Each objCell In objCopyRange
        i = i + 1
        If SameColor(Color, CellToColor(objCell)) Then
            Set objSameRange = UnionRange(objSameRange, objDestRange(i))
        Else
            Set objDiffRange = UnionRange(objDiffRange, objDestRange(i))
        End If
    Next
    
    '更新対象セルを高速化のためにクリア
    If GetTmpControl("C3").State Then
        '貼付け先領域全体をクリア
        Call ClearRange(objDestRange)
    Else
        '対象外のセルは更新しない時
        If GetTmpControl("C1").State Then
            '同じ色のセルをクリア
            Call ClearRange(objSameRange)
        Else
            '違う色のセルをクリア
            Call ClearRange(objDiffRange)
        End If
    End If
    
    '透明セルの設定
    If GetTmpControl("C3").State Then
        If GetTmpControl("C1").State Then
            '違う色のセルを透明化
            Call ColorToCell(objDiffRange, OleColorToARGB(&HFFFFFF, 0), False)
        Else
            '同じ色のセルを透明化
            Call ColorToCell(objSameRange, OleColorToARGB(&HFFFFFF, 0), False)
        End If
    End If
    
    'カラーの設定
    If GetTmpControl("C1").State Then
        '同じ色のセルを設定
        Call ColorToCell(objSameRange, Color, False)
    Else
        '違う色のセルを設定
        i = 0
        For Each objCell In objCopyRange
            i = i + 1
            If Not SameColor(Color, CellToColor(objCell)) Then
                Call ColorToCell(objDestRange(i), CellToColor(objCell), False)
            End If
        Next
    End If
End Sub


'*****************************************************************************
'[概要] 塗潰し
'[引数] なし
'[戻値] なし
'*****************************************************************************
Public Sub 塗潰し()
On Error GoTo ErrHandle
    Dim objSelection As Range
    If CheckSelection <> E_Range Then
        Call MsgBox("対象の範囲を選択してから実行してください")
        Exit Sub
    End If
    Set objSelection = Selection
    If objSelection.Rows.Count = Rows.Count Or objSelection.Columns.Count = Columns.Count Then
        Call MsgBox("すべての行または列の選択時は実行出来ません")
        Exit Sub
    End If
    If objSelection.Count = 1 Or objSelection.Areas.Count > 1 Then
        Call MsgBox("対象の範囲を選択してから実行してください")
        Exit Sub
    End If
    
    Dim objStartCell As Range
    Set objStartCell = SelectCell("塗潰し開始セルを選択してください", ActiveCell)
    If objStartCell Is Nothing Then
        Exit Sub
    End If
    If Intersect(objSelection, objStartCell) Is Nothing Then
        Call MsgBox("セルが、対象範囲の内側にありません")
        Exit Sub
    End If
    
    Dim objColorCell As Range
    Set objColorCell = SelectCell("塗潰し色のセルを選択してください", ActiveCell)
    If objColorCell Is Nothing Then
        Exit Sub
    End If
    
    Dim img As New CImage
    Call img.GetPixelsFromRange(objSelection)
    Call img.Fill(objStartCell.Column - objSelection.Column + 1, _
                  objStartCell.Row - objSelection.Row + 1, _
                  objColorCell)
    Application.ScreenUpdating = False
    Call SaveUndoInfo(objSelection)
    Call img.SetPixelsToRange(objSelection)
    Call SetOnUndo("塗潰し")
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
    If GetTmpControl("C4").State Or GetTmpControl("C5").State Or _
       GetTmpControl("C6").State Or GetTmpControl("C7").State Then
    Else
        Call MsgBox("RGBおよびアルファ値のいずれもチェックされていません")
        Exit Sub
    End If
    If Selection.Rows.Count = Rows.Count Or Selection.Columns.Count = Columns.Count Then
        Call MsgBox("すべての行または列の選択時は実行出来ません")
        Exit Sub
    End If

    If GetKeyState(vbKeyControl) < 0 Then
        lngUp = lngUp * 10
    End If
    
    '選択範囲の重複を排除
    Dim objSelection As Range
    Set objSelection = ReSelectRange(Selection)
    
    'RGBαの増減値
    Dim UpDown(1 To 4) As Long
    If GetTmpControl("C4").State Then
        UpDown(1) = lngUp
    End If
    If GetTmpControl("C5").State Then
        UpDown(2) = lngUp
    End If
    If GetTmpControl("C6").State Then
        UpDown(3) = lngUp
    End If
    If GetTmpControl("C7").State Then
        UpDown(4) = lngUp
    End If
    
    Dim objCell As Range
    Dim ARGB As TRGBQuad
    Application.ScreenUpdating = False
    Call SaveUndoInfo(Selection, "色調整")
    For Each objCell In objSelection
        ARGB = AdjustColor(CellToColor(objCell), UpDown(1), UpDown(2), UpDown(3), UpDown(4))
        Call ColorToCell(objCell, ARGB, True)
    Next
    Call Selection.Select
    Call SetOnUndo("色調整")
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[概要] 色のAHSLの値を増減させる
'[引数] 1:増加、-1:減少
'[戻値] なし
'*****************************************************************************
Public Sub HSL増減(ByVal lngUp As Long, ByVal lngType As Long)
On Error GoTo ErrHandle
    If CheckSelection <> E_Range Then Exit Sub
    If Selection.Rows.Count = Rows.Count Or Selection.Columns.Count = Columns.Count Then
        Call MsgBox("すべての行または列の選択時は実行出来ません")
        Exit Sub
    End If

    '増減値
    If GetKeyState(vbKeyControl) < 0 Then
        lngUp = lngUp * 10
    End If
    
    Dim H As Long
    Dim S As Long
    Dim L As Long
    Dim strUndo As String
    Select Case lngType
    Case 1 '色相
        H = lngUp
        strUndo = "色彩"
    Case 2 '彩度
        S = lngUp
        strUndo = "彩度"
    Case 3 '明度
        L = lngUp
        strUndo = "明度"
    End Select
    
    '選択範囲の重複を排除
    Dim objSelection As Range
    Set objSelection = ReSelectRange(Selection)
    
    Dim objCell As Range
    Dim ARGB As TRGBQuad
    Application.ScreenUpdating = False
    Call SaveUndoInfo(Selection, strUndo)
    For Each objCell In objSelection
        ARGB = UpDownHSL(CellToColor(objCell), H, S, L)
        Call ColorToCell(objCell, ARGB, True)
    Next
    Call Selection.Select
    Call SetOnUndo(strUndo)
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
'[概要] 色を数値化
'[引数] True:RGBA(16進8桁),False:RGB(16進6桁)
'[戻値] なし
'*****************************************************************************
Public Sub 色を数値化(ByVal blnAlpha As Boolean)
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
    Dim strValue As String
    For Each objCell In objSelection
        If blnAlpha Then
            strValue = Cell2RGBA(objCell)
        Else
            strValue = Cell2RGB(objCell)
        End If
        If strValue = "0" Then
            objCell.Value = 0
        Else
            objCell.Value = "'" & strValue
        End If
    Next
    
    'フォントの色と網掛けを標準に戻す
    With objSelection
        .Font.ColorIndex = xlAutomatic
        .Interior.Pattern = xlSolid
    End With
    Call Selection.Select
    Call SetOnUndo("色を数値化")
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub


'*****************************************************************************
'[概要] 数値から色を設定
'[引数] なし
'[戻値] なし
'*****************************************************************************
Public Sub 数値から色を設定()
On Error GoTo ErrHandle
    If CheckSelection <> E_Range Then Exit Sub
    If Selection.Rows.Count = Rows.Count Or Selection.Columns.Count = Columns.Count Then
        Call MsgBox("すべての行または列の選択時は実行出来ません")
        Exit Sub
    End If
        
    '値の入力されたセルのみ対象
    Dim objSelection As Range
    If Selection.Count <> 1 Then
        Dim objCells(1 To 3) As Range
        With Selection
            On Error Resume Next
            Set objCells(1) = .SpecialCells(xlCellTypeConstants)
            Set objCells(2) = .SpecialCells(xlCellTypeFormulas)
            On Error GoTo 0
        End With
        Set objSelection = UnionRange(objCells(1), objCells(2))
        If objSelection Is Nothing Then Exit Sub
    Else
        Set objSelection = Selection
    End If
    
    '対象セルを取得
    Dim objRange As Range
    Dim objZero As Range
    Dim objCell As Range
    Dim vValue  As Variant
    For Each objCell In objSelection
        vValue = objCell.Value
        If IsNumeric(vValue) Then
            If vValue = 0 Then
                Set objZero = UnionRange(objZero, objCell)
            End If
        ElseIf Left(vValue, 1) = "$" Then
            If Len(vValue) = 7 Or Len(vValue) = 9 Then
                If IsNumeric("&H" & Mid(vValue, 2)) Then
                   Set objRange = UnionRange(objRange, objCell)
                End If
            End If
        End If
    Next
    If (objRange Is Nothing) And (objZero Is Nothing) Then Exit Sub
        
    Application.ScreenUpdating = False
    Call SaveUndoInfo(Selection)
    
    '透明色以外
    If Not (objRange Is Nothing) Then
        '高速化のため書式をクリア
        With objRange
            .Interior.Pattern = xlNone
            .Font.Color = xlAutomatic
        End With
        
        Dim RGBA As TRGBQuad
        For Each objCell In objRange
            vValue = objCell.Value
            With RGBA
                .Red = "&H" & Mid(vValue, 2, 2)
                .Green = "&H" & Mid(vValue, 4, 2)
                .Blue = "&H" & Mid(vValue, 6, 2)
                If Len(vValue) = 9 Then
                    '8桁の時
                    .Alpha = "&H" & Mid(vValue, 8, 2)
                Else
                    '6桁の時、不透明
                    .Alpha = 255
                End If
            End With
            Call ColorToCell(objCell, RGBA, True)
        Next
    End If
    
    '透明色
    If Not (objZero Is Nothing) Then
        Call ColorToCell(objZero, OleColorToARGB(&HFFFFFF, 0), True)
    End If
    Call Selection.Select
    Call SetOnUndo("数値から色を設定")
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[概要] アルファ値を表示
'[引数] なし
'[戻値] なし
'*****************************************************************************
Public Sub アルファ値を表示()
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
    
On Error Resume Next
    Dim Alpha As Byte
    Dim vValue As String
    Dim objCell As Range
    For Each objCell In objSelection
        With objCell.Interior
            Select Case .ColorIndex
            Case xlNone, xlAutomatic
                '透明
                Alpha = 0
            Case Else
                '不透明
                Alpha = 255
                '半透明かどうか
                If .Pattern = xlGray8 Then
                    vValue = objCell.Value
                    If IsNumeric(vValue) Then
                        If 0 <= CLng(vValue) And CLng(vValue) <= 255 Then
                            'セルに入力された数値がアルファ値
                            Alpha = vValue
                        End If
                    End If
                End If
            End Select
        End With
        objCell.Value = Alpha
    Next
On Error GoTo ErrHandle
    
    'フォントの色を標準に戻す
    With objSelection.Font
        .ColorIndex = xlAutomatic
    End With
    Call SetOnUndo("α値表示")
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[概要] 数値からアルファ値を設定
'[引数] なし
'[戻値] なし
'*****************************************************************************
Public Sub 数値からアルファ値を設定()
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
    
On Error Resume Next
    '対象セルを取得
    Dim objRange As Range
    Dim objZero As Range
    Dim obj255 As Range
    Dim objCell As Range
    Dim vValue  As Variant
    For Each objCell In objSelection
        vValue = objCell.Value
        If IsNumeric(vValue) And vValue <> "" Then
            Select Case vValue
            Case 0 To 255
                Select Case objCell.Interior.ColorIndex
                Case xlNone, xlAutomatic
                    Set objZero = UnionRange(objZero, objCell)
                Case Else
                    Select Case vValue
                    Case 0
                        Set objZero = UnionRange(objZero, objCell)
                    Case 255
                        Set obj255 = UnionRange(obj255, objCell)
                    Case 1 To 254
                        Set objRange = UnionRange(objRange, objCell)
                    End Select
                End Select
            End Select
        End If
    Next
    If (objRange Is Nothing) And (objZero Is Nothing) And (obj255 Is Nothing) Then Exit Sub
        
    Application.ScreenUpdating = False
    Call SaveUndoInfo(Selection)
    
    '透明色
    If Not (objZero Is Nothing) Then
        Call ColorToCell(objZero, OleColorToARGB(&HFFFFFF, 0), True)
    End If
    
    '不透明色
    If Not (obj255 Is Nothing) Then
        With obj255
            .Interior.Pattern = xlAutomatic
            .Font.Color = xlAutomatic
            .ClearContents
        End With
    End If

    '半透明
    If Not (objRange Is Nothing) Then
        With objRange.Interior
            .Pattern = xlGray8
            .PatternColor = &HFFFFFF '白
        End With
        
        For Each objCell In objRange
            With objCell.Interior
                objCell.Font.Color = .Color '文字を背景色と同じにする
            End With
        Next
    End If
    
    Call Selection.Select
    Call SetOnUndo("α値設定")
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

