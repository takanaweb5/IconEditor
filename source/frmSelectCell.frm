VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectCell 
   Caption         =   "セルの選択"
   ClientHeight    =   1752
   ClientLeft      =   36
   ClientTop       =   396
   ClientWidth     =   3972
   OleObjectBlob   =   "frmSelectCell.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmSelectCell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit

'OKされたかどうか
Public IsOK As Boolean

'*****************************************************************************
'[イベント] フォームロード時
'*****************************************************************************
Private Sub UserForm_Initialize()
    Call RefEdit.SetFocus
    IsOK = False
End Sub

'*****************************************************************************
'[イベント] OKボタン押下時
'*****************************************************************************
Private Sub cmdOK_Click()
    Call cmdOK.SetFocus
    Call Me.Hide
    IsOK = True
End Sub

'*****************************************************************************
'[イベント] キャンセルボタン押下時
'*****************************************************************************
Private Sub cmdCancel_Click()
    Call cmdCancel.SetFocus
    Call Me.Hide
    IsOK = False
End Sub

