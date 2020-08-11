Attribute VB_Name = "GeneralTools"
Option Explicit
Option Private Module

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

Private Const CF_BITMAP = 2

'選択タイプ
Public Enum ESelectionType
    E_Range
    E_Shape
    E_Non
    E_Other
End Enum

Public Const MAX_WIDTH = 256
Public Const MAX_HEIGHT = 256

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

