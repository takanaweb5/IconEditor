VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectCell 
   Caption         =   "�Z���̑I��"
   ClientHeight    =   1752
   ClientLeft      =   36
   ClientTop       =   396
   ClientWidth     =   3972
   OleObjectBlob   =   "frmSelectCell.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmSelectCell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit

'OK���ꂽ���ǂ���
Public IsOK As Boolean

'*****************************************************************************
'[�C�x���g] �t�H�[�����[�h��
'*****************************************************************************
Private Sub UserForm_Initialize()
    Call RefEdit.SetFocus
    IsOK = False
End Sub

'*****************************************************************************
'[�C�x���g] OK�{�^��������
'*****************************************************************************
Private Sub cmdOK_Click()
    Call cmdOK.SetFocus
    Call Me.Hide
    IsOK = True
End Sub

'*****************************************************************************
'[�C�x���g] �L�����Z���{�^��������
'*****************************************************************************
Private Sub cmdCancel_Click()
    Call cmdCancel.SetFocus
    Call Me.Hide
    IsOK = False
End Sub

