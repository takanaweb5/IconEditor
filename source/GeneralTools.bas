Attribute VB_Name = "GeneralTools"
Option Explicit

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
    Dim Size As Long
    Dim p As LongPtr
    Size = GlobalSize(hMem)
    p = GlobalLock(hMem)
    ReDim Data(1 To Size) As Byte
    Call CopyMemory(Data(1), ByVal p, Size)
    Call GlobalUnlock(hMem)
    Call CloseClipboard
    hMem = 0
    
    Dim strData As String
    Dim i As Long
    For i = 1 To Size
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
