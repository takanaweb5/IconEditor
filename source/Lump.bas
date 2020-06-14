Attribute VB_Name = "Lump"
Option Explicit
Option Private Module

'*****************************************************************************
'[概要] 一括実行シートを開く
'[引数] なし
'[戻値] なし
'*****************************************************************************
Public Sub 一括実行シートを開く()
    With Worksheets("一括実行")
        .Visible = True
        .Activate
        .Range("A1").Select
    End With
End Sub

'*****************************************************************************
'[概要] 一括読込
'[引数] なし
'[戻値] なし
'*****************************************************************************
Public Sub 一括読込_Click()
On Error GoTo ErrHandle
    Dim objRange As Range
    Dim y As Long

    With ActiveSheet.Cells(1, 1)
        Set objRange = ActiveSheet.Range(.Cells(2, 1), .End(xlDown))
    End With
    
    Dim img As New CImage
    For y = 1 To objRange.Rows.Count
        Call img.LoadImageFromFile(objRange.Cells(y, "C"))
        Call img.SetPixelsToRange(Range(objRange.Cells(y, "E")))
        objRange.Cells(y, "F").Value = img.Width & "x" & img.Height
    Next
    Call MsgBox("処理が正常に終了しました")
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[概要] 一括保存
'[引数] なし
'[戻値] なし
'*****************************************************************************
Public Sub 一括保存_Click()
On Error GoTo ErrHandle
    Dim objRange As Range
    Dim y As Long

    With ActiveSheet.Cells(1, 1)
        Set objRange = ActiveSheet.Range(.Cells(2, 1), .End(xlDown))
    End With
    
    Dim objIconRange As Range
    Dim WidthAndHeight As Variant
    Dim ColCnt As Long
    Dim RowCnt As Long
    Dim img As New CImage
    For y = 1 To objRange.Rows.Count
        WidthAndHeight = Split(objRange.Cells(y, "F"), "x")
        ColCnt = 0
        RowCnt = 0
        If UBound(WidthAndHeight) = 1 Then
            If IsNumeric(WidthAndHeight(0)) And IsNumeric(WidthAndHeight(1)) Then
                ColCnt = WidthAndHeight(0)
                RowCnt = WidthAndHeight(1)
            End If
        End If
        If ColCnt > 0 And RowCnt > 0 Then
        Else
            Call MsgBox("サイズを正しく設定してください" & vbCrLf & objRange.Cells(y, "B"))
            Exit Sub
        End If
        
        Set objIconRange = Range(objRange.Cells(y, "E")).Resize(RowCnt, ColCnt)
        Call img.GetPixelsFromRange(objIconRange)
        Call img.SaveImageToFile(objRange.Cells(y, "C"))
    Next
    Call MsgBox("処理が正常に終了しました")
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

