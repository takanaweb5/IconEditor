Attribute VB_Name = "Ribbon"
Option Explicit
Option Private Module

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
'Public FChecked(1 To 7) As Boolean
Private FSampleClick As Boolean

'*****************************************************************************
'[概要] IRibbonUIを保存するCommandBarを作成する
'       あわせて、リボンコントロールの状態を保存するCommandBarControlを作成する
'       モジュール変数に保存した場合は、リコンパイルやコードの強制停止で値が損なわれるため
'[引数] IRibbonUI
'[戻値] なし
'*****************************************************************************
Private Sub CreateTmpCommandBar(ByRef Ribbon As IRibbonUI)
    On Error Resume Next
    Call Application.CommandBars(ThisWorkbook.Name).Delete
    On Error GoTo 0
    
    Dim i As Long
    Dim objCmdBar As CommandBar
    Set objCmdBar = CommandBars.Add(ThisWorkbook.Name, Position:=msoBarPopup, Temporary:=True)
    With objCmdBar.Controls.Add(msoControlButton)
        .Tag = "RibbonUI" & ThisWorkbook.Name
        .Parameter = ObjPtr(Ribbon)
    End With
    
    'チェックボックスのクローンをテンポラリに作成
    For i = 1 To 7
        With objCmdBar.Controls.Add(msoControlButton)
            .Tag = "C" & i & ThisWorkbook.Name
            .State = False '初期設定はチェックなし
        End With
    Next
    
    'RGBボタンの初期値をチェックありに設定
    GetTmpControl("C4").State = True 'Redチェックボックス
    GetTmpControl("C5").State = True 'Gereenチェックボックス
    GetTmpControl("C6").State = True 'Blueチェックボックス
End Sub

'*****************************************************************************
'[概要] CommandBarからIRibbonUIを取得する
'[引数] なし
'[戻値] IRibbonUI
'*****************************************************************************
Private Function GetRibbonUI() As IRibbonUI
    Dim Pointer  As LongPtr
    With CommandBars.FindControl(, , "RibbonUI" & ThisWorkbook.Name)
        Pointer = .Parameter
    End With
    Dim obj As Object
    Call CopyMemory(obj, Pointer, Len(Pointer))
    Set GetRibbonUI = obj
End Function

'*****************************************************************************
'[イベント] onLoad
'*****************************************************************************
Sub onLoad(Ribbon As IRibbonUI)
    'リボンUIをテンポラリのコマンドバーに保存する
    '(モジュール変数に保存した場合は、例外やコードの強制停止で値が損なわれるため)
    Call CreateTmpCommandBar(Ribbon)
'    Set FRibbon = Ribbon
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
        returnedVal = (GetTmpControl("C1").State Or GetTmpControl("C2").State)
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
    Case 11, 21, 32, 61, 62, 63
        returnedVal = 1
    Case Else
        returnedVal = 0
    End Select
End Sub

'*****************************************************************************
'[イベント] getPressed
'*****************************************************************************
Sub getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = GetTmpControl(control.ID).State
End Sub

'*****************************************************************************
'[イベント] onCheckAction
'*****************************************************************************
Sub onCheckAction(control As IRibbonControl, pressed As Boolean)
    'チェック状態を保存
    GetTmpControl(control.ID).State = pressed
    
    Select Case control.ID
    Case "C1"
'        Application.EnableEvents = False
        '特定色・特定色以外のトグル
        GetTmpControl("C2").State = False
        GetTmpControl("C3").State = False
        Call GetRibbonUI.InvalidateControl("C2")
        
        '有効無効を切り替る
        Call GetRibbonUI.InvalidateControl("C3")
'        Application.EnableEvents = True
    Case "C2"
'        Application.EnableEvents = False
        '特定色・特定色以外のトグル
        GetTmpControl("C1").State = False
        GetTmpControl("C3").State = False
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
        Result(1) = "色の" & vbCrLf & "貼付"
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
        Result(1) = "クリア"
        Result(2) = "クリア"
        Result(3) = "選択されたセル範囲の色をクリアします"
    Case 32
        Result(1) = "透明色の強調"
        Result(2) = "透明色の強調"
        Result(3) = "選択されたセル範囲で" & vbCrLf & "透明色のセルを網掛けで強調し、" & vbCrLf & "以外ののセルは網掛けを消します"
    Case 33
        Result(1) = "塗潰し"
        Result(2) = "塗潰し"
        Result(3) = "選択されたセル範囲を対象に塗潰します"
    Case 34
        Result(1) = "色の置換"
        Result(2) = "色の置換"
        Result(3) = "選択されたセル範囲の色を置換します"
    Case 35
        Result(1) = ""
        Result(2) = "左右反転"
        Result(3) = "選択されたセル範囲を左右反転します"
    Case 36
        Result(1) = ""
        Result(2) = "上下反転"
        Result(3) = "選択されたセル範囲を上下反転します"
    Case 37
        Result(1) = ""
        Result(2) = "右に90°回転"
        Result(3) = "選択されたセル範囲を右に90°回転します"
    Case 38
        Result(1) = ""
        Result(2) = "左に90°回転"
        Result(3) = "選択されたセル範囲を左に90°回転します"
    
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
        Result(1) = "色→数値(RGB)"
        Result(2) = "色を数値化"
        Result(3) = "色を6桁(RGB)の16進数でセルに設定します" & vbLf & "※ただし透明色は0とします"
    Case 72
        Result(1) = "色→数値(RGBA)"
        Result(2) = "色を数値化"
        Result(3) = "色を8桁(RGBA)の16進数でセルに設定します" & vbLf & "※ただし透明色は0とします"
    Case 73
        Result(1) = "α→数値"
        Result(2) = "セルの状態からアルファ値を表示"
        Result(3) = "無色のセルには、0を表示します" & vbLf & _
                    "セルが網掛けでα値が入力されているセルは、α値を表示します" & vbLf & _
                    "不透明のセルは255を表示します"
    Case 74
        Result(1) = "セル関数"
        Result(2) = "セル関数"
        Result(3) = "セル関数の紹介です"
    Case 75
        Result(1) = "数値→色"
        Result(2) = "セルの値から色を設定"
        Result(3) = "6桁または8桁の16進数のセルの値を色に変換します" & vbLf & _
                    "上記に該当しないセルは、何もしません" & vbLf & _
                    "8桁の時は右2桁をアルファ値に設定します" & vbLf & _
                    "6桁の時はアルファ値(透明･半透明)を設定しません" & vbLf & _
                    "実行後はセルの値をクリアします"
    Case 76
        Result(1) = "数値→α"
        Result(2) = "セルの数値からアルファ値を設定"
        Result(3) = "0のセルは透明にします" & vbLf & _
                    "1～254のセルはアルファ値を設定します" & vbLf & _
                    "以外(255や空白)のセルは、不透明にします" & vbLf & _
                    "実行後はセルの値をクリアします"
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
        Set objImage = Getイメージ(Range("Resource!1:1"))
    Case 22
        strImage = "PasteAsPicture"
    Case 23
        strImage = "PasteAsEmbedded"
    Case 24
        strImage = "ObjectRotateRight90"
    Case 25
        strImage = "ObjectRotateLeft90"
    
    Case 31
        strImage = "BlackAndWhiteWhite"
    Case 32
        strImage = "ViewGridlines"
    Case 33
'        strImage = "FillStyle"
        Set objImage = Getイメージ(Range("Resource!8:8"))
    Case 34
        Set objImage = Getイメージ(Range("Resource!7:7"))
    Case 35
        strImage = "ObjectFlipHorizontal"
    Case 36
        strImage = "ObjectFlipVertical"
    Case 37
        strImage = "ObjectRotateRight90"
    Case 38
        strImage = "ObjectRotateLeft90"
    
    Case 41
'        strImage = "TableSelectCellInfoPath"
        Set objImage = Getイメージ(Range("Resource!2:2"))
    Case 42
        Set objImage = Getイメージ(Range("Resource!3:3"))
    Case 43
'        strImage = "SelectSheet"
        Set objImage = Getイメージ(Range("Resource!4:4"))
    
    Case 51
        Set objImage = Getイメージ(Range("Resource!9:9"))
    Case 52
        Set objImage = Getイメージ(Range("Resource!10:10"))
    Case 53
        Set objImage = Getイメージ(Range("Resource!11:11"))
    Case 54
        Set objImage = Getイメージ(Range("Resource!12:12"))
    Case 55
        Set objImage = Getイメージ(Range("Resource!13:13"))
    Case 56
        Set objImage = Getイメージ(Range("Resource!14:14"))
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
    
    
    Case 74
        strImage = "EditFormula"
    Case 71, 72, 75
'        strImage = "ColorFuchsia"
        Set objImage = Getイメージ(Range("Resource!5:5"))
    Case 73, 76
'        strImage = "NotebookColor1"
        Set objImage = Getイメージ(Range("Resource!6:6"))
'    Case 77, 78
'        strImage = "DataTypeCalculatedColumn"
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
'   Call GetRibbonUI.Invalidate
    
    Select Case Mid(control.ID, 2)
    Case 11
        Call ImageMso取得
    Case 12
        Call 画像読込
    Case 13
        Call 一括実行シートを開く
    Case 14
        Call 画像保存
    Case 15
        Call 一括実行シートを開く
    Case 16
        Call 一括実行シートを開く
    
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
        Call クリア
    Case 32
        Call 透明色強調
    Case 33
        Call 塗潰し
    Case 34
        Call 色の置換
    Case 35
        Call 左右反転
    Case 36
        Call 上下反転
    Case 37
        Call 回転(1, 90)
    Case 38
        Call 回転(1, -90)
    
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
        Call 色を数値化(False)
    Case 72
        Call 色を数値化(True)
    Case 73
        Call アルファ値を表示
    Case 74
        Application.ScreenUpdating = False
        Call Worksheets("初めに").Activate
        ActiveWindow.ScrollColumn = 1
        ActiveWindow.ScrollRow = 90
        Call Worksheets("初めに").Range("A90").Select
        Application.ScreenUpdating = True
    Case 75
        Call 数値から色を設定
    Case 76
        Call 数値からアルファ値を設定
    End Select
End Sub


'*****************************************************************************
'[概要] セルのデータからアイコンファイルを読込む
'[引数] バイナリファイルデータを取得する行(Rangeオブジェクト)
'[戻値] IPicture
'*****************************************************************************
Private Function Getイメージ(ByRef objRange As Range) As IPicture
On Error GoTo ErrHandle
    Dim img As New CImage
    Call img.LoadImageFromResource(objRange)
    Set Getイメージ = img.SetToIPicture
ErrHandle:
End Function

'*****************************************************************************
'[概要] リボンのコールバック関数を実行する(Debug用)
'[引数] なし
'[戻値] なし
'*****************************************************************************
Public Sub InvalidateRibbon()
    Call GetRibbonUI.Invalidate
End Sub

