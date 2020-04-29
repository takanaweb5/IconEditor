Attribute VB_Name = "Ribbon"
Option Explicit

Private FRibbon As IRibbonUI
Public FChecked(1 To 7) As Boolean
Private FSampleClick As Boolean

Sub onLoad(ribbon As IRibbonUI)
    Set FRibbon = ribbon
End Sub

Sub loadImage(imageID As String, ByRef returnedVal)
  returnedVal = imageID
End Sub

Sub getVisible(control As IRibbonControl, ByRef returnedVal)
    returnedVal = True
End Sub

Sub getEnabled(control As IRibbonControl, ByRef returnedVal)
    Select Case control.ID
    Case "C3"
        returnedVal = FChecked(1) Or FChecked(2)
    Case Else
        returnedVal = True
    End Select
End Sub

Sub getShowLabel(control As IRibbonControl, ByRef returnedVal)
    returnedVal = GetTips(control, 0)
End Sub

Sub getLabel(control As IRibbonControl, ByRef returnedVal)
    returnedVal = GetTips(control, 1)
End Sub

Sub getScreentip(control As IRibbonControl, ByRef returnedVal)
    returnedVal = GetTips(control, 2)
End Sub

Sub getSupertip(control As IRibbonControl, ByRef returnedVal)
    returnedVal = GetTips(control, 3)
End Sub

'*****************************************************************************
'[概要] LabelおよびScreenTipを設定します
'[引数] lngType「0:getShowLabel, 1:getLabel, 2:getScreentip, 3:getSupertip」
'[戻値] 設定値
'*****************************************************************************
Private Function GetTips(control As IRibbonControl, ByVal lngType As Long) As Variant
    ReDim Result(1 To 3) '1:getLabel, 2:getScreentip, 3:getSupertip
    Select Case Mid(control.ID, 2)
    Case 11
        Result(1) = "ImageMso"
        Result(2) = "ImageMsoから画像を取得"
        Result(3) = "ImageMsoを指定して選択されたセルの位置へ画像を読み込みます"
    Case 12
        Result(1) = "読込"
        Result(2) = "画像の読込"
        Result(3) = "選択されたセルの位置へ画像を読込みます"
    Case 13
        Result(1) = "一括読込"
        Result(2) = "画像の一括読込み"
        Result(3) = "一括読込み用のシートを開きます"
    Case 14
        Result(1) = "保存"
        Result(2) = "画像の保存"
        Result(3) = "選択されたセル範囲の画像を保存します"
    Case 15
        Result(1) = "一括保存"
        Result(2) = "画像の一括保存"
        Result(3) = "一括保存用のシートを開きます"
    Case 16
        Result(1) = "Clipbord画像保存"
        Result(2) = "クリップボードの画像を保存"
        Result(3) = "クリップボードの画像をファイルに保存します"
    Case 21
        Result(1) = "貼付"
        Result(2) = "貼付け"
        Result(3) = "セルの色情報のみ貼付けます"
    Case 22
        Result(1) = "画像を貼付"
        Result(2) = "画像を貼付け"
        Result(3) = "クリップボードの画像を貼付けます"
    Case 23
        Result(1) = "Shapeを貼付"
        Result(2) = "Shapeを貼付け"
        Result(3) = "オートシェイプの画像を貼付けます"
    Case 24
        Result(1) = "90°"
        Result(2) = "右に90°回転して貼付け"
        Result(3) = "領域を右に90°回転して貼付けます"
    Case 25
        Result(1) = "-90°"
        Result(2) = "左に90°回転して貼付け"
        Result(3) = "領域を左に90°回転して貼付けます"
    Case 31
        Result(1) = ""
        Result(2) = "左右反転"
        Result(3) = "選択されたセル範囲を左右反転します"
    Case 32
        Result(1) = ""
        Result(2) = "上下反転"
        Result(3) = "選択されたセル範囲を上下反転します"
    Case 33
        Result(1) = ""
        Result(2) = "右に90°回転"
        Result(3) = "選択されたセル範囲を右に90°回転します"
    Case 34
        Result(1) = ""
        Result(2) = "左に90°回転"
        Result(3) = "選択されたセル範囲を左に90°回転します"
    Case 35
        Result(1) = "色の置換"
        Result(2) = "色の置換"
        Result(3) = "選択されたセル範囲の色を置換します"
    Case 36
        Result(1) = "塗潰し"
        Result(2) = "塗潰し"
        Result(3) = "選択されたセルの範囲を対象に塗潰します"
    
    Case 61
        Result(1) = "同じ色"
        Result(2) = "同じ色のセルの選択"
        Result(3) = "同じ色のセルを選択します"
    Case 62
        Result(1) = "違う色"
        Result(2) = "違う色のセルの選択"
        Result(3) = "違う色のセルを選択します"
    Case 63
        Result(1) = "反転等"
        Result(2) = "選択領域の反転など"
        Result(3) = "選択領域の反転や一部除外などが行えます"
    Case 71
        Result(1) = ""
        Result(2) = "RGB各色の数値およびアルファ値を減少させます"
        Result(3) = "選択した範囲のRGB各色の数値(0～255)を16(Ctrl押下時は1)減少させます"
    Case 72
        Result(1) = ""
        Result(2) = "RGB各色の数値およびアルファ値を増加させます"
        Result(3) = "選択した範囲のRGB各色の数値(0～255)を16(Ctrl押下時は1)増加させます"
    
    Case 41 To 46
        Select Case Mid(control.ID, 2)
        Case 41, 44
            Result(1) = "例1"
        Case 42, 45
            Result(1) = "例2"
        Case 43, 46
            Result(1) = "例3"
        End Select
        Select Case Mid(control.ID, 2)
        Case 41 To 43
            Result(2) = "サンプル表示(大)"
        Case Else
            Result(2) = "サンプル表示(小)"
        End Select
        Result(3) = "選択されたセル範囲の画像を表示します" & vbLf & _
                    "あわせてクリップボードにも画像をコピーします"
    Case 51
        Result(1) = "透明色の強調"
        Result(2) = "透明色の強調"
        Result(3) = "選択されたセル範囲の透明(半透明)色のセルを網掛けで強調します"
    Case 52
        Result(1) = "色→数値"
        Result(2) = "色を数値化"
        Result(3) = "色(RGB)を16進数でセルに設定します"
    Case 53
        Result(1) = "色→数値(α)"
        Result(2) = "色をアルファチャンネル付きで数値化"
        Result(3) = "色(ARGB)を16進数でセルに設定します"
    Case 54
        Result(1) = "数値→色"
        Result(2) = "セルの数値から色を設定"
        Result(3) = "セルの数値を色に変換します" & vbLf & _
                      "10進数・16進数のいずれも対応しています" & vbLf & _
                      "16進数で6桁以下の時はすべて不透明色に設定します" & vbLf & _
                      "10進数の時は、0は透明色と黒の判別がつかないので透明色とします"
    Case 55
        Result(1) = "数値→α"
        Result(2) = "セルの数値からアルファ値を設定"
        Result(3) = "セルの値が(0～255)の時、半透明度(アルファ値)を設定します" & vbLf & _
                      "0は完全透明" & vbLf & _
                      "255は完全不透明です"
    Case 56
        Result(1) = "16進→10進"
        Result(2) = "16進数→10進数"
        Result(3) = "セルの値が16進数の時、10進数に変換します"
    Case 57
        Result(1) = "10進→16進"
        Result(2) = "10進数→16進数"
        Result(3) = "セルの値が10進数の時、16進数に変換します"
    Case Else
        Result(1) = ""
    End Select
    
    If lngType = 0 Then
        GetTips = (Result(1) <> "")
    Else
        GetTips = Result(lngType)
    End If
End Function

Sub getShowImage(control As IRibbonControl, ByRef returnedVal)
    Call GetImages(control, 0, returnedVal)
End Sub

Sub getImage(control As IRibbonControl, ByRef returnedVal)
    Call GetImages(control, 1, returnedVal)
End Sub

Sub getSize(control As IRibbonControl, ByRef returnedVal)
    Call GetImages(control, 2, returnedVal)
End Sub

'*****************************************************************************
'[概要] Imageを設定します
'[引数] lngType「0:getShowImage, 1:getImage, 2:getSize」
'[戻値] Result
'*****************************************************************************
Private Sub GetImages(control As IRibbonControl, ByVal lngType As Long, ByRef Result)
    Dim strImage As String
    Dim lngSize  As Long
    Dim objImage As IPictureDisp

    Select Case Mid(control.ID, 2)
    Case 11
        strImage = "PictureReset"
        lngSize = 1 'large
    Case 12
        strImage = "FileOpen"
    Case 13
        strImage = "NewO12FilesTool"
    Case 14
        strImage = "FileSave"
    Case 15
        strImage = "SaveAll"
    Case 16
        strImage = "ObjectPictureFill"
    Case 21
        Set objImage = Getイメージ(Range("Icons!B2:AG33"))
        lngSize = 1 'large
    Case 22
        strImage = "PasteAsPicture"
    Case 23
        strImage = "PasteAsEmbedded"
    Case 24
        strImage = "ObjectRotateRight90"
    Case 25
        strImage = "ObjectRotateLeft90"
    Case 31
        strImage = "ObjectFlipHorizontal"
    Case 32
        strImage = "ObjectFlipVertical"
    Case 33
        strImage = "ObjectRotateRight90"
    Case 34
        strImage = "ObjectRotateLeft90"
    Case 35
        Set objImage = Getイメージ(Range("Icons!BP35:CU66"))
    Case 36
'        strImage = "FillStyle"
        Set objImage = Getイメージ(Range("Icons!CW35:EB66"))
    Case 61
'        strImage = "TableSelectCellInfoPath"
        Set objImage = Getイメージ(Range("Icons!AI2:BN33"))
    Case 62
        Set objImage = Getイメージ(Range("Icons!BP2:CU33"))
    Case 63
'        strImage = "SelectSheet"
        Set objImage = Getイメージ(Range("Icons!CW2:EB33"))
    Case 71
        strImage = "CatalogMergeGoToPreviousRecord"
    Case 72
        strImage = "CatalogMergeGoToNextRecord"
    Case 41 To 46
        If FSampleClick Then
            Set objImage = Getサンプル画像()
            FSampleClick = False
        Else
            strImage = "TentativeAcceptInvitation"
        End If
        If Mid(control.ID, 2) <= 43 Then
            lngSize = 1 'large
        Else
            lngSize = 0 'normal
        End If
    Case 51
        strImage = "ViewGridlines"
        lngSize = 1 'large
    Case 52
        strImage = "_1"
    Case 53
        strImage = "_2"
    Case 54
'        strImage = "ColorFuchsia"
        Set objImage = Getイメージ(Range("Icons!B35:AG66"))
    Case 55
'        strImage = "NotebookColor1"
        Set objImage = Getイメージ(Range("Icons!AI35:BN66"))
    Case 56, 57
        strImage = "DataTypeCalculatedColumn"
    Case Else
'        strimage = "BlackAndWhiteWhite"
    End Select

    Select Case lngType
    Case 0
        If (strImage = "") And (objImage Is Nothing) Then
            Result = False
        Else
            Result = True
        End If
    Case 1
        If strImage = "" Then
            Set Result = objImage
        Else
            Result = strImage
        End If
    Case 2
        Result = lngSize
    End Select
End Sub

Sub getPressed(control As IRibbonControl, ByRef returnedVal)
    Select Case Mid(control.ID, 2)
    Case 4 To 6
        returnedVal = True
        FChecked(Mid(control.ID, 2)) = True
    Case Else
        returnedVal = False
    End Select
End Sub

Sub onAction(control As IRibbonControl)
    Select Case Mid(control.ID, 2)
    Case 11
        Call ImageMso取得
    Case 12
        Call 画像読込
    Case 13
        Call FRibbon.Invalidate
    Case 14
        Call 画像保存
    Case 15
        Call FRibbon.Invalidate
    Case 16
        Call Clipbord画像保存
    Case 21
        Call 貼付け
    Case 22
        Call Clipbord画像読込
    Case 23
        Call Shape読込
    Case 24
        Call 回転(2, 90)
    Case 25
        Call 回転(2, -90)
    Case 31
        Call 左右反転
    Case 32
        Call 上下反転
    Case 33
        Call 回転(1, 90)
    Case 34
        Call 回転(1, -90)
    Case 35
        Call 色の置換
    Case 36
        Call 塗潰し
    Case 61
        Call 同色選択(True)
    Case 62
        Call 同色選択(False)
    Case 63
        Call 選択反転等
    Case 71
        Call 色増減(-1)
    Case 72
        Call 色増減(1)
    Case 41 To 46
        If CheckSelection <> E_Range Then Exit Sub
        Call Clipbord画像設定
        FSampleClick = True
        Call FRibbon.InvalidateControl(control.ID)
    Case 51
        Call 透明色強調
    End Select
End Sub

Sub onCheckAction(control As IRibbonControl, pressed As Boolean)
    Dim ID As Long
    ID = Mid(control.ID, 2)
    
    'チェック状態を保存
    FChecked(ID) = pressed
    Select Case ID
    Case 1
'        Application.EnableEvents = False
        '特定色・特定色以外のトグル
        FChecked(2) = False
        Call FRibbon.InvalidateControl("C2")
        
        '有効無効を切り替る
        Call FRibbon.InvalidateControl("C3")
'        Application.EnableEvents = True
    Case 2
'        Application.EnableEvents = False
        '特定色・特定色以外のトグル
        FChecked(1) = False
        Call FRibbon.InvalidateControl("C1")
        
        '有効無効を切り替る
        Call FRibbon.InvalidateControl("C3")
'        Application.EnableEvents = True
    End Select
End Sub

Private Function Getイメージ(ByRef objRange As Range) As IPictureDisp
On Error GoTo ErrHandle
    Dim img As New CImage
    Call img.GetPixelsFromRange(objRange)
    Set Getイメージ = img.SetToIPicture
ErrHandle:
End Function
