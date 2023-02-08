Attribute VB_Name = "GeneralTools"
Option Explicit
Option Private Module

Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As Long
Public Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long
Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Public Declare PtrSafe Function IsZoomed Lib "user32" (ByVal hwnd As LongPtr) As Long
Public Declare PtrSafe Function GetSystemMenu Lib "user32" (ByVal hwnd As LongPtr, ByVal bRevert As Long) As LongPtr
Public Declare PtrSafe Function EnableMenuItem Lib "user32.dll" (ByVal hMenu As LongPtr, ByVal uIDEnableItem As Long, ByVal uEnable As Long) As Long
Public Declare PtrSafe Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As LongPtr
Public Declare PtrSafe Function GetExitCodeProcess Lib "kernel32.dll" (ByVal hProcess As LongPtr, lpExitCode As Long) As Long
Public Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
Public Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As Long
Public Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hDC As LongPtr) As Long

Public Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
Public Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Public Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
Public Declare PtrSafe Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
Public Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Public Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As Long
Public Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As LongPtr
Public Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As LongPtr
Public Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As Long
Public Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As Long
Public Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Public Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Public Declare PtrSafe Function GetKeyState Lib "user32" (ByVal lngVirtKey As Long) As Integer
Public Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long

Public Declare PtrSafe Function ImmGetContext Lib "imm32.dll" (ByVal hwnd As LongPtr) As LongPtr
Public Declare PtrSafe Function ImmSetOpenStatus Lib "imm32.dll" (ByVal himc As LongPtr, ByVal B As Long) As Long
Public Declare PtrSafe Function ImmReleaseContext Lib "imm32.dll" (ByVal hwnd As LongPtr, ByVal himc As LongPtr) As Long

Public Declare PtrSafe Function CopyEnhMetaFile Lib "gdi32" Alias "CopyEnhMetaFileA" (ByVal hemfSrc As LongPtr, ByVal lpszFile As String) As LongPtr
Public Declare PtrSafe Function DeleteEnhMetaFile Lib "gdi32" (ByVal hEmf As LongPtr) As Long

Public Const CF_BITMAP = 2
Public Const CF_ENHMETAFILE = 14

'選択タイプ
Public Enum ESelectionType
    E_Range
    E_Shape
    E_Non
    E_Other
End Enum

Public Const MAX_WIDTH = 256
Public Const MAX_HEIGHT = 256

' 定数の定義
Public Const SC_CLOSE = 61536
Public Const SC_SIZE = &HF000&
Public Const MF_BYCOMMAND = 0
Public Const MF_GRAYED = 1
Public Const LOGPIXELSX = 88
Public Const LOGPIXELSY = 90

'ソート用構造体
Public Type TSortArray
    Key1  As Long
    Key2  As Long
    Key3  As Long
End Type

'*****************************************************************************
'[概要] 選択されているかオブジェクトの種類を判定する
'[引数] なし
'[戻値] Range、Shape、その他　のいずれか
'*****************************************************************************
Public Function CheckSelection() As ESelectionType
On Error GoTo ErrHandle
    If ActiveWorkbook Is Nothing Then
        CheckSelection = E_Non
        Exit Function
    End If
    
    If Selection Is Nothing Then
        CheckSelection = E_Other
        Exit Function
    End If
    
    If TypeOf Selection Is Range Then
        CheckSelection = E_Range
    ElseIf TypeOf Selection.ShapeRange Is ShapeRange Then
        CheckSelection = E_Shape
    Else
        CheckSelection = E_Other
    End If
Exit Function
ErrHandle:
    CheckSelection = E_Other
End Function

'*****************************************************************************
'[概要] コピー対象のRangeを取得する
'[引数] なし
'[戻値] コピー対象のRange
'*****************************************************************************
Public Function GetCopyRange() As Range
    If OpenClipboard(0) = 0 Then Exit Function
    Dim hMem As LongPtr
    hMem = GetClipboardData(RegisterClipboardFormat("Link"))
    If hMem = 0 Then
        Call CloseClipboard
        Exit Function
    End If
     
On Error GoTo ErrHandle
    Dim size As Long
    Dim p As LongPtr
    size = GlobalSize(hMem)
    p = GlobalLock(hMem)
    ReDim Data(1 To size) As Byte
    Call CopyMemory(Data(1), ByVal p, size)
    Call GlobalUnlock(hMem)
    Call CloseClipboard
    hMem = 0
    
    Dim strData As String
    Dim i As Long
    For i = 1 To size
        If Data(i) = 0 Then
            Data(i) = Asc("/") 'シート名にもファイル名にも使えない文字
        End If
    Next i
    strData = StrConv(Data, vbUnicode)
    
    Dim objRegExp As Object
    Set objRegExp = CreateObject("VBScript.RegExp")
    objRegExp.Global = False
    objRegExp.Pattern = "Excel\/.*\[(.+)\](.+)\/(.+)\/\/$"
    If Not objRegExp.Test(strData) Then Exit Function
    With objRegExp.Execute(strData)(0)
        Dim strRange As String
        strRange = Application.ConvertFormula(.SubMatches(2), xlR1C1, xlA1)
        Set GetCopyRange = Workbooks(.SubMatches(0)).Worksheets(.SubMatches(1)).Range(strRange)
    End With
    Application.CutCopyMode = False
    Exit Function
ErrHandle:
    If hMem <> 0 Then Call CloseClipboard
End Function

'*****************************************************************************
'[概要] クリップボードにBitmap形式がコピーされているかどうか
'[引数] なし
'[戻値] True:Bitmap形式あり
'*****************************************************************************
Public Function ClipboardHasBitmap() As Boolean
    ClipboardHasBitmap = (IsClipboardFormatAvailable(CF_BITMAP) <> 0)
End Function

'*****************************************************************************
'[概要] フォームを表示してセルを選択させる
'[引数] 表示するメッセージ、objCurrentCell：初期選択させるセル
'[戻値] 選択されたセル（キャンセル時はNothing）
'*****************************************************************************
Public Function SelectCell(ByVal strMsg As String, ByRef objCurrentCell As Range) As Range
    Dim strCell As String
    'フォームを表示
    With frmSelectCell
        .Label.Caption = strMsg
        Call objCurrentCell.Worksheet.Activate
        .RefEdit.Text = objCurrentCell.AddressLocal
        Call .Show
        If .IsOK = True Then
            strCell = .RefEdit
        End If
    End With
    Call Unload(frmSelectCell)
    If strCell <> "" Then
        Set SelectCell = Range(strCell)
        If SelectCell.Address = SelectCell.Cells(1, 1).MergeArea.Address Then
            Set SelectCell = SelectCell.Cells(1, 1)
        End If
    End If
End Function

'*****************************************************************************
'[概要] 拡張子の取得
'[引数] ファイルパス
'[戻値] 拡張子(大文字)
'*****************************************************************************
Public Function GetFileExtension(ByVal strFilename As String) As String
    With CreateObject("Scripting.FileSystemObject")
        GetFileExtension = UCase(.GetExtensionName(strFilename))
    End With
End Function

'*****************************************************************************
'[概要] 領域と領域の重なる領域を取得する
'[引数] 対象領域(Nothingも可)
'[戻値] objRange1 ∩ objRange2
'*****************************************************************************
Public Function IntersectRange(ByRef objRange1 As Range, ByRef objRange2 As Range) As Range
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
Public Function UnionRange(ByRef objRange1 As Range, ByRef objRange2 As Range) As Range
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
Public Function MinusRange(ByRef objRange1 As Range, ByRef objRange2 As Range) As Range
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
Public Function ReSelectRange(ByRef objRange As Range) As Range
    Set ReSelectRange = objRange.Areas(1)
    
    Dim i As Long
    For i = 2 To objRange.Areas.Count
        Set ReSelectRange = Union(ReSelectRange, ReSelectRange(MinusRange(objRange.Areas(i), ReSelectRange)))
    Next
End Function

'*****************************************************************************
'[概要] 領域が一致するか判定
'[引数] 対象領域アドレス
'[戻値] True:一致
'*****************************************************************************
'Public Function IsSameRange(ByRef strRange1 As String, ByRef strRange2 As String) As Boolean
'    If strRange1 = "" Or strRange2 = "" Then
'        Exit Function
'    End If
'
'    Dim objRange1 As Range
'    Dim objRange2 As Range
'    Set objRange1 = AddressToRange(strRange1)
'    Set objRange2 = AddressToRange(strRange2)
'    IsSameRange = MinusRange(objRange1, objRange2) Is Nothing
'    If IsSameRange Then
'        IsSameRange = MinusRange(objRange2, objRange1) Is Nothing
'    End If
'End Function

'*****************************************************************************
'[概要] Rangeのアドレスを取得する(255字以上に対応するため)
'[引数] Range
'[戻値] 例：A1:C3/E5/F1:G5
'*****************************************************************************
Public Function RangeToAddress(ByRef objRange As Range) As String
    ReDim Address(1 To objRange.Areas.Count)
    Dim i As Long
    For i = 1 To objRange.Areas.Count
        Address(i) = objRange.Areas(i).Address(False, False)
    Next
    RangeToAddress = Join(Address, "/")
End Function

'*****************************************************************************
'[概要] RangeToAddressの結果からRangeを取得する
'[引数] 例：A1:C3/E5/F1:G5
'[戻値] Range
'*****************************************************************************
Public Function AddressToRange(ByVal strAddress As String) As Range
    Dim Address As Variant
    Address = Split(strAddress, "/")
    Dim i As Long
    For i = LBound(Address) To UBound(Address)
        Set AddressToRange = UnionRange(AddressToRange, Range(Address(i)))
    Next
End Function

'*****************************************************************************
'[概要] セルの色をクリアする
'[引数] 対象領域
'[戻値] なし
'*****************************************************************************
Public Function ClearRange(ByRef objRange As Range)
    If objRange Is Nothing Then Exit Function
    With objRange
        .Interior.Pattern = xlNone
        .Font.ColorIndex = xlAutomatic
        .ClearContents
    End With
End Function

'*****************************************************************************
'[概要] テンポラリのCommandBarControlを取得する
'[引数] Controlを識別するID（リボンコントロールのID）
'[戻値] CommandBarControl
'*****************************************************************************
Public Function GetTmpControl(ByVal strID As String) As CommandBarControl
    Set GetTmpControl = CommandBars.FindControl(, , strID & ThisWorkbook.Name)
End Function

'*****************************************************************************
'[概要] バイナリファイルをセルに読込む
'[引数] 読込むファイル名, バイナリファイルを読込む行(Rangeオブジェクト)
'[戻値] なし
'*****************************************************************************
Public Sub LoadResourceFromFile(ByVal strFilename As String, ByRef objRow As Range)
    'A列はファイル名とする
    objRow.Cells(1, 1).Value = Dir(strFilename)
    
    'ファイルサイズの配列を作成
    ReDim Data(1 To FileLen(strFilename)) As Byte

    Dim File As Integer
    File = FreeFile()
    Open strFilename For Binary Access Read As #File
    Get #File, , Data
    Close #File
    
    Dim x As Long
    For x = 1 To UBound(Data)
        objRow.Cells(1, x + 1) = Data(x)
    Next
End Sub

'*****************************************************************************
'[概要] セルのデータをバイナリファイルを書込む
'[引数] 書込むファイル名, バイナリファイルデータを取得する行(Rangeオブジェクト)
'[戻値] なし
'*****************************************************************************
Public Sub SaveResourceToFile(ByVal strFilename As String, ByRef objRow As Range)
    'ファイルサイズの配列を作成
    ReDim Data(1 To objRow.Cells(1, 1).End(xlToRight).Column - 1) As Byte
    Dim x As Long
    For x = 1 To UBound(Data)
         Data(x) = objRow.Cells(1, x + 1)
    Next
    
    Dim File As Integer
    File = FreeFile()
    Open strFilename For Binary Access Write As #File
    Put #File, , Data
    Close #File
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
                GetUndoStr = Trim(.List(1))
            End If
        End If
    End With
End Function

'*****************************************************************************
'[概要] 変更対象の図形の中で回転しているものをグループ化する
'[引数] グループ化前の図形
'[戻値] グループ化後の図形
'*****************************************************************************
Public Function GroupSelection(ByRef objShapes As ShapeRange) As ShapeRange
    Dim i            As Long
    Dim objShape     As Shape
    Dim btePlacement As Byte
    ReDim blnRotation(1 To objShapes.Count) As Boolean
    ReDim lngIDArray(1 To objShapes.Count) As Variant
    
    '図形の数だけループ
    For i = 1 To objShapes.Count
        Set objShape = objShapes(i)
        lngIDArray(i) = objShape.ID
        
        Select Case objShape.Rotation
        Case 90, 270, 180
            blnRotation(i) = True
        End Select
    Next

    '図形の数だけループ
    For i = 1 To objShapes.Count
        If blnRotation(i) = True Then
            Set objShape = GetShapeFromID(lngIDArray(i))
            btePlacement = objShape.Placement
            'サイズと位置が同一のクローンを作成しグループ化する
            With objShape.Duplicate
                .Top = objShape.Top
                .Left = objShape.Left
                If objShape.Top < 0 Then
                    '図形が回転して座標がマイナスになった時ゼロになるため補正する
                    Call .IncrementTop(objShape.Top)
                End If
                If objShape.Left < 0 Then
                    '図形が回転して座標がマイナスになった時ゼロになるため補正する
                    Call .IncrementLeft(objShape.Left)
                End If
                
                '透明にする
                .Fill.Visible = msoFalse
                .Line.Visible = msoFalse
                With GetShapeRangeFromID(Array(.ID, objShape.ID)).Group
                    .AlternativeText = "EL_TemporaryGroup" & i
                    .Placement = btePlacement
                    lngIDArray(i) = .ID
                End With
            End With
        End If
    Next
    
    Set GroupSelection = GetShapeRangeFromID(lngIDArray)
End Function

'*****************************************************************************
'[概要] 変更対象の図形の中でグループ化したものを元に戻す
'[引数] グループ解除前の図形
'[戻値] グループ解除後の図形
'*****************************************************************************
Public Function UnGroupSelection(ByRef objGroups As ShapeRange) As ShapeRange
    Dim i            As Long
    Dim btePlacement As Byte
    Dim objShape     As Shape
    ReDim blnRotation(1 To objGroups.Count) As Boolean
    ReDim lngIDArray(1 To objGroups.Count) As Variant
    
    '図形の数だけループ
    For i = 1 To objGroups.Count
        Set objShape = objGroups(i)
        lngIDArray(i) = objShape.ID
        
        If Left$(objShape.AlternativeText, 17) = "EL_TemporaryGroup" Then
            blnRotation(i) = True
        End If
    Next

    '図形の数だけループ
    For i = 1 To objGroups.Count
        If blnRotation(i) = True Then
            Set objShape = GetShapeFromID(lngIDArray(i))
            btePlacement = objShape.Placement
            With objShape.Ungroup
                .Item(1).Placement = btePlacement
                Call .Item(2).Delete
                lngIDArray(i) = .Item(1).ID
            End With
        End If
    Next i
    
    Set UnGroupSelection = GetShapeRangeFromID(lngIDArray)
End Function

'*****************************************************************************
'[概要] ShapeオブジェクトのIDからShapeオブジェクトを取得
'[引数] ID
'[戻値] Shapeオブジェクト
'*****************************************************************************
Private Function GetShapeFromID(ByVal lngID As Long) As Shape
    Dim j As Long
    Dim lngIndex As Long
        
    For j = 1 To ActiveSheet.Shapes.Count
        If ActiveSheet.Shapes(j).ID = lngID Then
            lngIndex = j
            Exit For
        End If
    Next j
    
    Set GetShapeFromID = ActiveSheet.Shapes.Range(j).Item(1)
End Function

'*****************************************************************************
'[概要] ShpesオブジェクトのIDからShapeRangeオブジェクトを取得
'[引数] IDの配列
'[戻値] ShapeRangeオブジェクト
'*****************************************************************************
Public Function GetShapeRangeFromID(ByRef lngID As Variant) As ShapeRange
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
'[概要] Shapeの四方に最も近いセル範囲を取得する
'[引数] Shapeオブジェクト
'[戻値] セル範囲
'*****************************************************************************
Public Function GetNearlyRange(ByRef objShape As Shape) As Range
    Dim objTopLeft     As Range
    Dim objBottomRight As Range
    Set objTopLeft = objShape.TopLeftCell
    Set objBottomRight = objShape.BottomRightCell
    
    '上の位置と高さを設定
    If objShape.Height = 0 Then
        With objTopLeft
            If .Top + .Height / 2 < objShape.Top Then
                Set objTopLeft = Cells(.Row + 1, .Column)
                Set objBottomRight = Cells(.Row + 1, objBottomRight.Column)
            End If
        End With
    Else
        '下のセルの再設定
        With objBottomRight
            If .Top = objShape.Top + objShape.Height Then
                Set objBottomRight = Cells(.Row - 1, .Column)
            End If
        End With
            
        '上端の再設定
        With objTopLeft
            If .Top + .Height / 2 < objShape.Top Then
                If .Row + 1 <= objBottomRight.Row Then
                    Set objTopLeft = Cells(.Row + 1, .Column)
                End If
            End If
        End With
                
        '下端の再設定
        With objBottomRight
            If .Top + .Height / 2 > objShape.Top + objShape.Height Then
                If .Row - 1 >= objTopLeft.Row Then
                    Set objBottomRight = Cells(.Row - 1, .Column)
                End If
            End If
        End With
    End If
    
    '左の位置と幅を設定
    If objShape.Width = 0 Then
        With objTopLeft
            If .Left + .Width / 2 < objShape.Left Then
                Set objTopLeft = Cells(.Row, .Column + 1)
                Set objBottomRight = Cells(objBottomRight.Row, .Column + 1)
            End If
        End With
    Else
        '右のセルの再設定
        With objBottomRight
            If .Left = objShape.Left + objShape.Width Then
                Set objBottomRight = Cells(.Row, .Column - 1)
            End If
        End With
    
        '左端の再設定
        With objTopLeft
            If .Left + .Width / 2 < objShape.Left Then
                If .Column + 1 <= objBottomRight.Column Then
                    Set objTopLeft = Cells(.Row, .Column + 1)
                End If
            End If
        End With
                
        '右端の再設定
        With objBottomRight
            If .Left + .Width / 2 > objShape.Left + objShape.Width Then
                If .Column - 1 >= objTopLeft.Column Then
                    Set objBottomRight = Cells(.Row, .Column - 1)
                End If
            End If
        End With
    End If
    
    Set GetNearlyRange = Range(objTopLeft, objBottomRight)
End Function

'*****************************************************************************
'[概要] DPIの変換率を取得する ※72(ExcelのデフォルトのDPI)/画面のDPI
'[引数] なし
'[戻値] なし
'*****************************************************************************
Public Function DPIRatio() As Double
    DPIRatio = 72 / GetDPI()
End Function

'*****************************************************************************
'[概要] DPIを取得する
'[引数] なし
'[戻値] DPI ※標準は96
'*****************************************************************************
Public Function GetDPI() As Long
    Dim DC As LongPtr
    DC = GetDC(0)
    GetDPI = GetDeviceCaps(DC, LOGPIXELSX)
    Call ReleaseDC(0, DC)
End Function

'*****************************************************************************
'[概要] SortArray配列をソートする
'[引数] Sort対象配列
'[戻値] なし
'*****************************************************************************
Public Sub SortArray(ByRef SortArray() As TSortArray)
    'バブルソート
    Dim i As Long
    Dim j As Long
    Dim Swap As TSortArray
    For i = UBound(SortArray) To 1 Step -1
        For j = 1 To i - 1
            If CompareValue(SortArray(j), SortArray(j + 1)) Then
                Swap = SortArray(j)
                SortArray(j) = SortArray(j + 1)
                SortArray(j + 1) = Swap
            End If
        Next j
    Next i
End Sub

'*****************************************************************************
'[概要] 大小比較を行う
'[引数] 比較対象
'[戻値] True: SortArray1 > SortArray2
'*****************************************************************************
Private Function CompareValue(ByRef SortArray1 As TSortArray, ByRef SortArray2 As TSortArray) As Boolean
    If SortArray1.Key1 = SortArray2.Key1 Then
        CompareValue = (SortArray1.Key2 > SortArray2.Key2)
    Else
        CompareValue = (SortArray1.Key1 > SortArray2.Key1)
    End If
End Function

'*****************************************************************************
'[概要] 入力の位置の左横の枠線の位置を取得(単位ピクセル)
'[引数] lngPos:位置(単位ピクセル)
'       objColumn: lngPosを含む列
'[戻値] 図形の左側の枠線の位置(単位ピクセル)
'*****************************************************************************
Public Function GetLeftGrid(ByVal lngPos As Long, ByRef objColumn As Range) As Long
    Dim i       As Long
    Dim lngLeft As Long
    
    If lngPos <= Round(Columns(2).Left / DPIRatio) Then
        GetLeftGrid = 0
        Exit Function
    End If
    
    For i = objColumn.Column To 1 Step -1
        lngLeft = Round(GetWidth(Range(Columns(1), Columns(i - 1))) / DPIRatio)
        If lngLeft < lngPos Then
            GetLeftGrid = lngLeft
            Exit Function
        End If
    Next
End Function

'*****************************************************************************
'[概要] 入力の位置の右横の枠線の位置を取得(単位ピクセル)
'[引数] lngPos:位置(単位ピクセル)
'       objColumn: lngPosを含む列
'[戻値] 図形の右側の枠線の位置(単位ピクセル)
'*****************************************************************************
Public Function GetRightGrid(ByVal lngPos As Long, ByRef objColumn As Range) As Long
    Dim i        As Long
    Dim lngRight As Long
    
    If lngPos >= Round(GetWidth(Range(Columns(1), Columns(Columns.Count - 1))) / DPIRatio) Then
        GetRightGrid = Round(GetWidth(Columns) / DPIRatio)
        Exit Function
    End If
    
    For i = objColumn.Column + 1 To Columns.Count
        lngRight = Round(GetWidth(Range(Columns(1), Columns(i - 1))) / DPIRatio)
        If lngRight > lngPos Then
            GetRightGrid = lngRight
            Exit Function
        End If
    Next
End Function

'*****************************************************************************
'[概要] 入力の位置の上の枠線の位置を取得(単位ピクセル)
'[引数] lngPos:位置(単位ピクセル)
'       objRow: lngPosを含む行
'[戻値] 図形の上側の枠線の位置(単位ピクセル)
'*****************************************************************************
Public Function GetTopGrid(ByVal lngPos As Long, ByRef objRow As Range) As Long
    Dim i      As Long
    Dim lngTop As Long
    
    If lngPos <= Round(Rows(2).Top / DPIRatio) Then
        GetTopGrid = 0
        Exit Function
    End If
    
    For i = objRow.Row To 1 Step -1
        lngTop = Round(Rows(i).Top / DPIRatio)
        If lngTop < lngPos Then
            GetTopGrid = lngTop
            Exit Function
        End If
    Next
End Function

'*****************************************************************************
'[概要] 入力の位置の下の枠線の位置を取得(単位ピクセル)
'[引数] lngPos:位置(単位ピクセル)
'       objRow: lngPosを含む行
'[戻値] 図形の下側の枠線の位置(単位ピクセル)
'*****************************************************************************
Public Function GetBottomGrid(ByVal lngPos As Long, ByRef objRow As Range) As Long
    Dim i         As Long
    Dim lngBottom As Long
    Dim lngMax    As Long
    
    lngMax = Round((Rows(Rows.Count).Top + Rows(Rows.Count).Height) / DPIRatio)
    
    If lngPos >= Round(Rows(Rows.Count).Top / DPIRatio) Then
        GetBottomGrid = lngMax
        Exit Function
    End If
    
    For i = objRow.Row + 1 To Rows.Count
        lngBottom = Round(Rows(i).Top / DPIRatio)
        If lngBottom > lngPos Then
            GetBottomGrid = lngBottom
            Exit Function
        End If
    Next
End Function

'*****************************************************************************
'[概要] 選択エリアの幅を取得
'       Width/Leftプロパティは32767以上の幅を計算出来ないため
'[引数] 幅を取得するエリア
'[戻値] 幅(Widthプロパティ)
'*****************************************************************************
Private Function GetWidth(ByRef objRange As Range) As Double
    Dim lngCount   As Long
    Dim lngHalf    As Long
    Dim MaxWidth   As Double '幅の最大値

    MaxWidth = 32767 * DPIRatio
    If objRange.Width < MaxWidth Then
        GetWidth = objRange.Width
    Else
        With objRange
            '前半＋後半の幅を合計
            lngCount = .Columns.Count
            lngHalf = lngCount / 2
            GetWidth = GetWidth(Range(.Columns(1), .Columns(lngHalf))) + _
                       GetWidth(Range(.Columns(lngHalf + 1), .Columns(lngCount)))
        End With
    End If
End Function


