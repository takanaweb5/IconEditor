VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmUnSelect 
   Caption         =   "選択してください"
   ClientHeight    =   2592
   ClientLeft      =   108
   ClientTop       =   336
   ClientWidth     =   4668
   OleObjectBlob   =   "frmUnSelect.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmUnSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Declare PtrSafe Function GetKeyState Lib "user32" (ByVal lngVirtKey As Long) As Integer

'領域の取消し画面のモード
Public Enum EUnselectMode
    E_Unselect  '取消し
    E_Reverse   '反転
    E_Union     '追加
    E_Intersect '絞り込み
End Enum

Private lngReferenceStyle As Long
Private strSelectBefore As String
Private blnCheck As Boolean

Private strLastSheet   As String '前回の領域の復元用
Private strLastAddress As String '前回の領域の復元用

'*****************************************************************************
'[概要] 各種マウス操作時、RefEditを有効にさせる
'*****************************************************************************
Private Sub UserForm_MouseMove(ByVal button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    On Error Resume Next
    RefEdit.SetFocus
End Sub
Private Sub Frame_Click()
    On Error Resume Next
    RefEdit.SetFocus
End Sub
Private Sub Frame_MouseMove(ByVal button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    On Error Resume Next
    RefEdit.SetFocus
End Sub
Private Sub lblTitle_Click()
    On Error Resume Next
    RefEdit.SetFocus
End Sub
Private Sub lblTitle_MouseMove(ByVal button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    On Error Resume Next
    RefEdit.SetFocus
End Sub
Private Sub cmdOK_MouseMove(ByVal button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    On Error Resume Next
    RefEdit.SetFocus
End Sub
Private Sub cmdCancel_MouseMove(ByVal button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    On Error Resume Next
    RefEdit.SetFocus
End Sub

'*****************************************************************************
'[概要] RefEditで領域選択時にアドレスが255文字を超えた時の対応
'*****************************************************************************
Private Sub RefEdit_Change()
    '[Ctrl]Keyが押下されていれば、選択領域を次々と追加している時
    If GetKeyState(vbKeyControl) < 0 Then
        'アドレスが255文字を超えてリセットされてしまった時
        If strSelectBefore <> "" Then
            If Range(RefEdit.Value).Areas.Count <= 1 And _
               Range(strSelectBefore).Areas.Count > 1 Then
                Call MsgBox("これ以上は選択出来ません", vbExclamation)
                RefEdit.Value = strSelectBefore
                Call cmdOK.SetFocus
                Call RefEdit.SetFocus
                Exit Sub
            End If
        End If
    End If
    strSelectBefore = RefEdit.Value
End Sub

'*****************************************************************************
'[概要] RefEditで領域選択中にCtrlキーを離した時の対応
'*****************************************************************************
Private Sub RefEdit_KeyUp(KeyCode As Integer, ByVal Shift As Integer)
    If KeyCode = vbKeyControl Then
        RefEdit.Value = strSelectBefore
        Call cmdOK.SetFocus
        Call RefEdit.SetFocus
    End If
End Sub

'*****************************************************************************
'[概要] 反転チェック時
'*****************************************************************************
Private Sub chkReverse_Change()
    If blnCheck = True Then
        Exit Sub
    End If
    Call ChangeMode(E_Reverse)
End Sub

'*****************************************************************************
'[概要] 絞り込みチェック時
'*****************************************************************************
Private Sub chkIntersect_Change()
    If blnCheck = True Then
        Exit Sub
    End If
    If chkIntersect.Value = True Then
        Call ChangeMode(E_Intersect)
    Else
        Call ChangeMode(E_Reverse)
    End If
End Sub

'*****************************************************************************
'[概要] 追加チェック時
'*****************************************************************************
Private Sub chkUnion_Change()
    If blnCheck = True Then
        Exit Sub
    End If
    If chkUnion.Value = True Then
        Call ChangeMode(E_Union)
    Else
        Call ChangeMode(E_Reverse)
    End If
End Sub

'*****************************************************************************
'[概要] 取消しチェック時
'*****************************************************************************
Private Sub chkUnselect_Change()
    If blnCheck = True Then
        Exit Sub
    End If
    If chkUnselect.Value = True Then
        Call ChangeMode(E_Unselect)
    Else
        Call ChangeMode(E_Reverse)
    End If
End Sub

'*****************************************************************************
'[概要]  ｢反転｣･｢絞り込み｣･｢取消し｣のモードを変更する
'[引数] モードタイプ
'[戻値] なし
'*****************************************************************************
Private Sub ChangeMode(ByVal enmModeType As EUnselectMode)
    Select Case enmModeType
    Case E_Reverse
        lblTitle.Caption = "マウスで選択を反転させたい領域を選択してください"
    Case E_Intersect
        lblTitle.Caption = "マウスで選択を絞り込みたい領域を選択してください"
    Case E_Union
        lblTitle.Caption = "マウスで選択を追加したい領域を選択してください"
    Case E_Unselect
        lblTitle.Caption = "マウスで選択を取消させたい領域を選択してください"
    End Select
    
    blnCheck = True
    chkReverse.Value = (enmModeType = E_Reverse)
    chkIntersect.Value = (enmModeType = E_Intersect)
    chkUnion.Value = (enmModeType = E_Union)
    chkUnselect.Value = (enmModeType = E_Unselect)
 
    blnCheck = False
    RefEdit.Enabled = True
    Call RefEdit.SetFocus
End Sub
    
'*****************************************************************************
'[概要] フォームロード時
'*****************************************************************************
Private Sub UserForm_Initialize()
    lngReferenceStyle = Application.ReferenceStyle
    Application.ReferenceStyle = xlA1

    'RefEditを隠す
    RefEdit.Top = RefEdit.Top + 100
    
    '呼び元に通知する
    FFormLoad = True
    
    Call ChangeMode(E_Reverse)
End Sub

'*****************************************************************************
'[概要] フォームアンロード時
'*****************************************************************************
Private Sub UserForm_Terminate()
    Application.ReferenceStyle = lngReferenceStyle
    '呼び元に通知する
    FFormLoad = False
End Sub

'*****************************************************************************
'[概要] 前回の領域の復元ボタン押下時
'*****************************************************************************
Private Sub cmdLastSelect_Click()
    Call Range(strLastAddress).Select
End Sub

'*****************************************************************************
'[概要] ＯＫボタン押下時
'*****************************************************************************
Private Sub cmdOK_Click()
    Call cmdOK.SetFocus
    Me.Hide
End Sub

'*****************************************************************************
'[概要] キャンセルボタン押下時
'*****************************************************************************
Private Sub cmdCancel_Click()
    Call Unload(Me)
End Sub

'*****************************************************************************
'[概要] 直前のコマンド実行時に選択された領域のアドレスを保存する
'[引数] 直前の領域の情報
'[戻値] なし
'*****************************************************************************
Public Sub SetLastSelect(ByVal strSheetName As String, ByVal strAddress As String)
    strLastSheet = strSheetName
    strLastAddress = strAddress
    
    If strLastAddress = "" Or ActiveSheet.Name <> strLastSheet Then
        cmdLastSelect.Enabled = False
    End If
End Sub

'*****************************************************************************
'[概要] 選択された領域
'[引数] なし
'*****************************************************************************
Public Property Get SelectRange() As Range
    Dim objRange  As Range
    Dim strAddress As String
    
    For Each objRange In Range(RefEdit.Value).Areas
        strAddress = strAddress & objRange.Address(False, False) & ","
    Next
    
    '末尾のカンマを削除
    Set SelectRange = Range(Left$(strAddress, Len(strAddress) - 1))
End Property
'*****************************************************************************
'[概要] 選択モード
'[引数] なし
'*****************************************************************************
Public Property Get Mode() As EUnselectMode
    Select Case (True)
    Case chkReverse.Value
        Mode = E_Reverse
    Case chkIntersect.Value
        Mode = E_Intersect
    Case chkUnion.Value
        Mode = E_Union
    Case Else
        Mode = E_Unselect
    End Select
End Property
