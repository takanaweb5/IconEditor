Attribute VB_Name = "Lump"
Option Explicit
Option Private Module

'*****************************************************************************
'[�T�v] �ꊇ���s�V�[�g���J��
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub �ꊇ���s�V�[�g���J��()
    With Worksheets("�ꊇ���s")
        .Visible = True
        .Activate
        .Range("A1").Select
    End With
End Sub

'*****************************************************************************
'[�T�v] �ꊇ�Ǎ�
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub �ꊇ�Ǎ�_Click()
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
    Call MsgBox("����������ɏI�����܂���")
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[�T�v] �ꊇ�ۑ�
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub �ꊇ�ۑ�_Click()
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
            Call MsgBox("�T�C�Y�𐳂����ݒ肵�Ă�������" & vbCrLf & objRange.Cells(y, "B"))
            Exit Sub
        End If
        
        Set objIconRange = Range(objRange.Cells(y, "E")).Resize(RowCnt, ColCnt)
        Call img.GetPixelsFromRange(objIconRange)
        Call img.SaveImageToFile(objRange.Cells(y, "C"))
    Next
    Call MsgBox("����������ɏI�����܂���")
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

