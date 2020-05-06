Attribute VB_Name = "UndoInfo"
Option Explicit

Private Const UndoSheetName = "Undo"
Private FRange     As Range   'Undoの対象領域
Private FSelection As String  '選択領域のアドレス

'*****************************************************************************
'[概要] Undo情報を保存する
'[引数] Undoする領域
'[戻値] なし
'*****************************************************************************
Public Sub SaveUndoInfo(ByRef objSelection As Range, Optional strCommand As String = "")
    If strCommand <> "" Then
        '色の調整コマンド等が連打されている時
        If strCommand = Left(GetUndoStr(), Len(strCommand)) Then
            If objSelection.Address(False, False) = FSelection Then
                Exit Sub
            End If
        End If
    End If
    
    Dim objSheet As Worksheet
    Set objSheet = ThisWorkbook.Worksheets(UndoSheetName)
    
    Call ClearUndoSheet
    FSelection = objSelection.Address(False, False)
    Set FRange = GetCanvas(objSelection)
    Call FRange.Copy(objSheet.Range(FRange.Address))
End Sub

'*****************************************************************************
'[概要] Undo用の領域全体を取得
'[引数] Undoする領域
'[戻値] Undoする領域が複数の時、すべてを包括する領域を取得
'*****************************************************************************
Private Function GetCanvas(ByRef objSelection As Range) As Range
    Dim lngRow(1 To 2) As Long '1:最小値,2:最大値
    Dim lngCol(1 To 2) As Long '1:最小値,2:最大値

    '最大値を初期設定
    lngRow(1) = Rows.Count
    lngCol(1) = Columns.Count
    
    Dim objArea As Range
    For Each objArea In objSelection.Areas
        '領域ごとの一番左上のセル
        With objArea.Cells(1)
            lngRow(1) = WorksheetFunction.Min(lngRow(1), .Row)
            lngCol(1) = WorksheetFunction.Min(lngCol(1), .Column)
        End With
        '領域ごとの一番右下のセル
        With objArea.Cells(objArea.Cells.Count)
            lngRow(2) = WorksheetFunction.MAX(lngRow(2), .Row)
            lngCol(2) = WorksheetFunction.MAX(lngCol(2), .Column)
        End With
    Next
    
    Dim objCell(1 To 2) As Range
    Set objCell(1) = objSelection.Worksheet.Cells(lngRow(1), lngCol(1))
    Set objCell(2) = objSelection.Worksheet.Cells(lngRow(2), lngCol(2))
    Set GetCanvas = Range(objCell(1), objCell(2))
End Function

'*****************************************************************************
'[概要] ApplicationオブジェクトのOnUndoイベントを設定
'[引数] Undoに表示するコマンド名
'[戻値] なし
'*****************************************************************************
Public Sub SetOnUndo(ByVal strCommand As String)
    Call Application.OnUndo(strCommand, "ExecUndo")
End Sub

'*****************************************************************************
'[概要] Undoを実行する
'[引数] なし
'[戻値] なし
'*****************************************************************************
Public Sub ExecUndo()
On Error GoTo Finalization
    Dim objSheet As Worksheet
    Set objSheet = ThisWorkbook.Worksheets(UndoSheetName)
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Call objSheet.Range(FRange.Address).Copy(FRange)
    Call FRange.Worksheet.Activate
    Call Range(FSelection).Select
    Set FRange = Nothing
    Call ClearUndoSheet
Finalization:
    Application.DisplayAlerts = True
End Sub

'*****************************************************************************
'[概要] ワークシートの中身をクリアする
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub ClearUndoSheet()
    Dim objSheet As Worksheet
    Set objSheet = ThisWorkbook.Worksheets(UndoSheetName)
    
    Dim objShape  As Shape
    For Each objShape In objSheet.Shapes
        Call objShape.Delete
    Next
    Call objSheet.Cells.Delete
    '最後のセルを修正する
    Call objSheet.Cells.Parent.UsedRange
End Sub

'*****************************************************************************
'[概要] Undoボタンの情報を取得する
'[引数] なし
'[戻値] UndoボタンのTooltipText
'*****************************************************************************
Public Function GetUndoStr() As String
    With CommandBars.FindControl(, 128) 'Undoボタン
        If .Enabled Then
            If .ListCount = 1 Then
                'Undoが1種類の時のUndoコマンド
                GetUndoStr = .TooltipText
            End If
        End If
    End With
End Function

