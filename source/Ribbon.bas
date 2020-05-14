Attribute VB_Name = "Ribbon"
Option Explicit

Private Const PAGE_READONLY = 2&
Private Const PAGE_READWRITE = 4&
Private Const FILE_MAP_WRITE = 2&
Private Const FILE_MAP_READ = 4&

Private Declare PtrSafe Function CreateFileMapping Lib "kernel32" Alias "CreateFileMappingW" (ByVal hFile As LongPtr, lpFileMappingAttributes As Any, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As String) As LongPtr
Private Declare PtrSafe Function OpenFileMapping Lib "kernel32" Alias "OpenFileMappingW" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal ptrToNameString As String) As LongPtr
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function MapViewOfFile Lib "kernel32" (ByVal hFileMappingObject As LongPtr, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As LongPtr
Private Declare PtrSafe Function UnmapViewOfFile Lib "kernel32" (ByVal lpBaseAddress As LongPtr) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)

'Private FRibbon As IRibbonUI '例外等が起きても値が損なわれないように共有メモリに変更
Public FChecked(1 To 7) As Boolean
Private FSampleClick As Boolean

'*****************************************************************************
'[概要] IRibbonUIを共有メモリに保存する
'[引数] IRibbonUI
'[戻値] なし
'*****************************************************************************
Private Sub SetRibbonUI(ByRef Ribbon As IRibbonUI)
    Dim hFileMap As LongPtr
    Dim pMap     As LongPtr
    Dim Pointer  As LongPtr

    'ハンドルをCloseすることはあきらめる
    hFileMap = CreateFileMapping(-1, ByVal 0&, PAGE_READWRITE, 0, Len(Pointer), ThisWorkbook.FullName)
'    hFileMap = OpenFileMapping(FILE_MAP_WRITE, False, ThisWorkbook.FullName)
    If hFileMap <> 0 Then
        pMap = MapViewOfFile(hFileMap, FILE_MAP_WRITE, 0, 0, 0)
        If pMap <> 0 Then
            Pointer = ObjPtr(Ribbon)
            Call CopyMemory(ByVal pMap, Pointer, Len(Pointer))
            Call UnmapViewOfFile(pMap)
        End If
    End If
'    Set FRibbon = Ribbon
End Sub

'*****************************************************************************
'[概要] IRibbonUIを共有メモリから取得する
'[引数] なし
'[戻値] IRibbonUI
'*****************************************************************************
Private Function GetRibbonUI() As IRibbonUI
    Dim hFileMap As LongPtr
    Dim pMap     As LongPtr
    Dim Pointer  As LongPtr

    hFileMap = OpenFileMapping(FILE_MAP_READ, False, ThisWorkbook.FullName)
    If hFileMap <> 0 Then
        pMap = MapViewOfFile(hFileMap, FILE_MAP_READ, 0, 0, 0)
        If pMap <> 0 Then
            Call CopyMemory(Pointer, ByVal pMap, Len(Pointer))
            Call UnmapViewOfFile(pMap)

            Dim obj As Object
            Call CopyMemory(obj, Pointer, Len(Pointer))
            Set GetRibbonUI = obj
        End If
        Call CloseHandle(hFileMap)
    End If
'    GetRibbonUI = FRibbon
End Function

'*****************************************************************************
'[イベント] onLoad
'*****************************************************************************
Sub onLoad(Ribbon As IRibbonUI)
    'リボンUIを共有メモリに保存する
    '(モジュール変数に保存した場合は、例外やコードのBreakで値が損なわれるため)
    Call SetRibbonUI(Ribbon)
End Sub

'*****************************************************************************
'[イベント] loadImage
'*****************************************************************************
Sub loadImage(imageID As String, ByRef returnedVal)
  returnedVal = imageID
End Sub

'*****************************************************************************
'[イベント] getVisible
'*****************************************************************************
Sub getVisible(control As IRibbonControl, ByRef returnedVal)
    returnedVal = True
End Sub

'*****************************************************************************
'[イベント] getEnabled
'*****************************************************************************
Sub getEnabled(control As IRibbonControl, ByRef returnedVal)
    Select Case control.ID
    Case "C3"
        returnedVal = FChecked(1) Or FChecked(2)
    Case Else
        returnedVal = True
    End Select
End Sub

'*****************************************************************************
'[イベント] getShowLabel
'*****************************************************************************
Sub getShowLabel(control As IRibbonControl, ByRef returnedVal)
    returnedVal = (GetTips(control, 1) <> "")
End Sub

'*****************************************************************************
'[イベント] getLabel
'*****************************************************************************
Sub getLabel(control As IRibbonControl, ByRef returnedVal)
    returnedVal = GetTips(control, 1)
End Sub

'*****************************************************************************
'[イベント] getScreentip
'*****************************************************************************
Sub getScreentip(control As IRibbonControl, ByRef returnedVal)
    returnedVal = GetTips(control, 2)
End Sub

'*****************************************************************************
'[イベント] getSupertip
'*****************************************************************************
Sub getSupertip(control As IRibbonControl, ByRef returnedVal)
    returnedVal = GetTips(control, 3)
End Sub

'*****************************************************************************
'[イベント] getShowImage
'*****************************************************************************
Sub getShowImage(control As IRibbonControl, ByRef returnedVal)
    Select Case Mid(control.ID, 2)
    Case Else
        returnedVal = True
    End Select
End Sub

'*****************************************************************************
'[イベント] getImage
'*****************************************************************************
Sub getImage(control As IRibbonControl, ByRef returnedVal)
    Call GetImages(control, returnedVal)
End Sub

'*****************************************************************************
'[イベント] getSize
'*****************************************************************************
Sub getSize(control As IRibbonControl, ByRef returnedVal)
    Select Case Mid(control.ID, 2)
    Case 11, 21, 37, 61, 62, 63, 71
        returnedVal = 1
    Case Else
        returnedVal = 0
    End Select
End Sub

'*****************************************************************************
'[イベント] getPressed
'*****************************************************************************
Sub getPressed(control As IRibbonControl, ByRef returnedVal)
    Select Case Mid(control.ID, 2)
    Case 4 To 6
        returnedVal = True
        FChecked(Mid(control.ID, 2)) = True
    Case Else
        returnedVal = False
    End Select
End Sub

'*****************************************************************************
'[イベント] onCheckAction
'*****************************************************************************
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
        Call GetRibbonUI.InvalidateControl("C2")
        
        '有効無効を切り替る
        Call GetRibbonUI.InvalidateControl("C3")
'        Application.EnableEvents = True
    Case 2
'        Application.EnableEvents = False
        '特定色・特定色以外のトグル
        FChecked(1) = False
        Call GetRibbonUI.InvalidateControl("C1")
        
        '有効無効を切り替る
        Call GetRibbonUI.InvalidateControl("C3")
'        Application.EnableEvents = True
    End Select
End Sub

'*****************************************************************************
'[概要] LabelおよびScreenTipを設定します
'[引数] lngType「1:getLabel, 2:getScreentip, 3:getSupertip」
'[戻値] 設定値
'*****************************************************************************
Private Function GetTips(control As IRibbonControl, ByVal lngType As Long) As String
    ReDim Result(1 To 3) '1:getLabel, 2:getScreentip, 3:getSupertip
    Select Case Mid(control.ID, 2)
    Case 11
        Result(1) = "ImageMso"
        Result(2) = "ImageMsoから画像を取得"
        Result(3) = "ImageMsoを指定して選択されたセル範囲へ画像を読み込みます" & vbCrLf & "単一セルの選択時はサイズを指定するダイアログが表示されます"
    Case 12
        Result(1) = "読込"
        Result(2) = "画像の読込"
        Result(3) = "選択されたセル範囲へ画像を読み込みます" & vbCrLf & "単一セルの選択時は元の画像のサイズで読み込みます"
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
        Result(3) = "クリップボードの画像を選択されたセル範囲へ貼付けます" & vbCrLf & "単一セルの選択時は元の画像のサイズで貼付けます"
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
        Result(3) = "選択されたセル範囲を対象に塗潰します"
    Case 37
        Result(1) = "クリア"
        Result(2) = "クリア"
        Result(3) = "選択されたセル範囲の色をクリアします"
    
    Case 41
        Result(1) = "同じ色"
        Result(2) = "同じ色のセルの選択"
        Result(3) = "同じ色のセルを選択します"
    Case 42
        Result(1) = "違う色"
        Result(2) = "違う色のセルの選択"
        Result(3) = "違う色のセルを選択します"
    Case 43
        Result(1) = "反転等"
        Result(2) = "選択領域の反転など"
        Result(3) = "選択領域の反転や一部除外などが行えます"
    Case 51, 52
        Result(1) = ""
        Result(2) = "明るさ"
        Result(3) = "選択した範囲の色の明るさ(0～255)を1単位(Ctrlで10単位)で変更します" & vbCrLf & "増加:黒→原色→白" & vbCrLf & "減少:黒←原色←白"
    Case 53, 54
        Result(1) = ""
        Result(2) = "彩やかさ"
        Result(3) = "選択した範囲の色の彩やかさ(0～255)を1単位(Ctrlで10単位)で変更します" & vbCrLf & "増加:灰色→純色" & vbCrLf & "減少:灰色←純色"
    Case 55, 56
        Result(1) = ""
        Result(2) = "色相"
        Result(3) = "選択した範囲の色相(0～360°)を1°単位(Ctrlで5°単位)で変化させます" & vbCrLf & "増加:赤→黄→緑→青→紫→赤" & vbCrLf & "減少:赤←黄←緑←青←紫←赤"
    Case 57
        Result(1) = ""
        Result(2) = "RGB各色の数値(0～255)およびアルファ値を減少させます"
        Result(3) = "選択した範囲のチェックしたRGB各色の数値(0～255)を1単位(Ctrlで10単位)で減少させます"
    Case 58
        Result(1) = ""
        Result(2) = "RGB各色の数値およびアルファ値を増加させます"
        Result(3) = "選択した範囲のチェックしたRGB各色の数値(0～255)を1単位(Ctrlで10単位)で増加させます"
    
    Case 61 To 66
        Select Case Mid(control.ID, 2)
        Case 61, 64
            Result(1) = "例1"
        Case 62, 65
            Result(1) = "例2"
        Case 63, 66
            Result(1) = "例3"
        End Select
        Select Case Mid(control.ID, 2)
        Case 61 To 63
            Result(2) = "サンプル表示(大)"
        Case Else
            Result(2) = "サンプル表示(小)"
        End Select
        Result(3) = "選択されたセル範囲の画像を表示します" & vbLf & _
                    "あわせてクリップボードにも画像をコピーします"
    Case 71
        Result(1) = "透明色の強調"
        Result(2) = "透明色の強調"
        Result(3) = "選択されたセル範囲の透明(半透明)色のセルを網掛けで強調します"
    Case 72
        Result(1) = "色→数値"
        Result(2) = "色を数値化"
        Result(3) = "色(RGB)を16進数でセルに設定します"
    Case 73
        Result(1) = "色→数値(α)"
        Result(2) = "色をアルファチャンネル付きで数値化"
        Result(3) = "色(ARGB)を16進数でセルに設定します"
    Case 74
        Result(1) = "数値→色"
        Result(2) = "セルの数値から色を設定"
        Result(3) = "セルの数値を色に変換します" & vbLf & _
                      "10進数・16進数のいずれも対応しています" & vbLf & _
                      "16進数で6桁以下の時はすべて不透明色に設定します" & vbLf & _
                      "10進数の時は、0は透明色と黒の判別がつかないので透明色とします"
    Case 75
        Result(1) = "数値→α"
        Result(2) = "セルの数値からアルファ値を設定"
        Result(3) = "セルの値が(0～255)の時、半透明度(アルファ値)を設定します" & vbLf & _
                      "0は完全透明" & vbLf & _
                      "255は完全不透明です"
    Case 76
        Result(1) = "16進→10進"
        Result(2) = "16進数→10進数"
        Result(3) = "セルの値が16進数の時、10進数に変換します"
    Case 77
        Result(1) = "10進→16進"
        Result(2) = "10進数→16進数"
        Result(3) = "セルの値が10進数の時、16進数に変換します"
    Case Else
        Result(1) = ""
    End Select
    
    GetTips = Result(lngType)
End Function

'*****************************************************************************
'[概要] Imageを設定します
'[引数] control
'[戻値] Result
'*****************************************************************************
Private Sub GetImages(control As IRibbonControl, ByRef Result)
    Dim strImage As String
    Dim objImage As IPictureDisp

    Select Case Mid(control.ID, 2)
    Case 11
        strImage = "PictureReset"
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
    Case 37
        strImage = "ViewGridlines"
    Case 36
'        strImage = "FillStyle"
        Set objImage = Getイメージ(Range("Icons!CW35:EB66"))
    Case 41
'        strImage = "TableSelectCellInfoPath"
        Set objImage = Getイメージ(Range("Icons!AI2:BN33"))
    Case 42
        Set objImage = Getイメージ(Range("Icons!BP2:CU33"))
    Case 43
'        strImage = "SelectSheet"
        Set objImage = Getイメージ(Range("Icons!CW2:EB33"))
    
    Case 51
        Set objImage = Getイメージ(Range("Icons!B68:AG99"))
    Case 52
        Set objImage = Getイメージ(Range("Icons!AI68:BN99"))
    
    Case 53
        Set objImage = Getイメージ(Range("Icons!B101:AG132"))
    Case 54
        Set objImage = Getイメージ(Range("Icons!AI101:BN132"))
    Case 55
        Set objImage = Getイメージ(Range("Icons!B134:AG165"))
    Case 56
        Set objImage = Getイメージ(Range("Icons!AI134:BN165"))
    
    Case 57
        strImage = "CatalogMergeGoToPreviousRecord"
    Case 58
        strImage = "CatalogMergeGoToNextRecord"
    Case 61 To 66
        If FSampleClick Then
            Set objImage = Getサンプル画像()
            FSampleClick = False
        Else
            strImage = "TentativeAcceptInvitation"
        End If
    Case 71
        strImage = "ViewGridlines"
    Case 72
        strImage = "_1"
    Case 73
        strImage = "_2"
    Case 74
'        strImage = "ColorFuchsia"
        Set objImage = Getイメージ(Range("Icons!B35:AG66"))
    Case 75
'        strImage = "NotebookColor1"
        Set objImage = Getイメージ(Range("Icons!AI35:BN66"))
    Case 76, 77
        strImage = "DataTypeCalculatedColumn"
    Case Else
'        strimage = "BlackAndWhiteWhite"
    End Select

    If strImage = "" Then
        Set Result = objImage
    Else
        Result = strImage
    End If
End Sub

'*****************************************************************************
'[イベント] onAction
'*****************************************************************************
Sub onAction(control As IRibbonControl)
    Select Case Mid(control.ID, 2)
    Case 11
        Call ImageMso取得
    Case 12
        Call 画像読込
    Case 13
        Call GetRibbonUI.Invalidate
    Case 14
        Call 画像保存
    Case 15
        Call GetRibbonUI.Invalidate
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
    Case 37
        Call クリア
    Case 41
        Call 同色選択(True)
    Case 42
        Call 同色選択(False)
    Case 43
        Call 選択反転等
    Case 51
        Call HSL増減(-1, 3)
    Case 52
        Call HSL増減(1, 3)
    Case 53
        Call HSL増減(-1, 2)
    Case 54
        Call HSL増減(1, 2)
    Case 55
        Call HSL増減(-1, 1)
    Case 56
        Call HSL増減(1, 1)
    Case 57
        Call 色増減(-1)
    Case 58
        Call 色増減(1)
    Case 61 To 66
        If CheckSelection <> E_Range Then Exit Sub
        Call Clipbord画像設定
        FSampleClick = True
        Call GetRibbonUI.InvalidateControl(control.ID)
    Case 71
        Call 透明色強調
    End Select
End Sub

'*****************************************************************************
'[イベント] loadImage
'*****************************************************************************
Private Function Getイメージ(ByRef objRange As Range) As IPicture
On Error GoTo ErrHandle
    Dim img As New CImage
    Call img.GetPixelsFromRange(objRange)
    Set Getイメージ = img.SetToIPicture
ErrHandle:
End Function
