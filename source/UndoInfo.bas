Attribute VB_Name = "UndoInfo"
Option Explicit

Private Const UndoSheetName = "Undo"
Private FRange     As Range   'Undo�̑Ώۗ̈�
Private FSelection As String  '�I��̈�̃A�h���X

'*****************************************************************************
'[�T�v] Undo����ۑ�����
'[����] Undo����̈�
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub SaveUndoInfo(ByRef objSelection As Range, Optional strCommand As String = "")
    If strCommand <> "" Then
        '�F�̒����R�}���h�����A�ł���Ă��鎞
        If strCommand = Left(GetUndoStr(), Len(strCommand)) Then
            If objSelection.Address(False, False) = FSelection Then
                Exit Sub
            End If
        End If
    End If
    
    Dim objSheet As Worksheet
    Set objSheet = ThisWorkbook.Worksheets(UndoSheetName)
    
    Call ClearUndoSheet
    FSelection = objSelection.Address(False, False)
    Set FRange = GetCanvas(objSelection)
    Call FRange.Copy(objSheet.Range(FRange.Address))
End Sub

'*****************************************************************************
'[�T�v] Undo�p�̗̈�S�̂��擾
'[����] Undo����̈�
'[�ߒl] Undo����̈悪�����̎��A���ׂĂ�����̈���擾
'*****************************************************************************
Private Function GetCanvas(ByRef objSelection As Range) As Range
    Dim lngRow(1 To 2) As Long '1:�ŏ��l,2:�ő�l
    Dim lngCol(1 To 2) As Long '1:�ŏ��l,2:�ő�l

    '�ő�l�������ݒ�
    lngRow(1) = Rows.Count
    lngCol(1) = Columns.Count
    
    Dim objArea As Range
    For Each objArea In objSelection.Areas
        '�̈悲�Ƃ̈�ԍ���̃Z��
        With objArea.Cells(1)
            lngRow(1) = WorksheetFunction.Min(lngRow(1), .Row)
            lngCol(1) = WorksheetFunction.Min(lngCol(1), .Column)
        End With
        '�̈悲�Ƃ̈�ԉE���̃Z��
        With objArea.Cells(objArea.Cells.Count)
            lngRow(2) = WorksheetFunction.MAX(lngRow(2), .Row)
            lngCol(2) = WorksheetFunction.MAX(lngCol(2), .Column)
        End With
    Next
    
    Dim objCell(1 To 2) As Range
    Set objCell(1) = objSelection.Worksheet.Cells(lngRow(1), lngCol(1))
    Set objCell(2) = objSelection.Worksheet.Cells(lngRow(2), lngCol(2))
    Set GetCanvas = Range(objCell(1), objCell(2))
End Function

'*****************************************************************************
'[�T�v] Application�I�u�W�F�N�g��OnUndo�C�x���g��ݒ�
'[����] Undo�ɕ\������R�}���h��
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub SetOnUndo(ByVal strCommand As String)
    Call Application.OnUndo(strCommand, "ExecUndo")
End Sub

'*****************************************************************************
'[�T�v] Undo�����s����
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub ExecUndo()
On Error GoTo Finalization
    Dim objSheet As Worksheet
    Set objSheet = ThisWorkbook.Worksheets(UndoSheetName)
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Call objSheet.Range(FRange.Address).Copy(FRange)
    Call FRange.Worksheet.Activate
    Call Range(FSelection).Select
    Set FRange = Nothing
    Call ClearUndoSheet
Finalization:
    Application.DisplayAlerts = True
End Sub

'*****************************************************************************
'[�T�v] ���[�N�V�[�g�̒��g���N���A����
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub ClearUndoSheet()
    Dim objSheet As Worksheet
    Set objSheet = ThisWorkbook.Worksheets(UndoSheetName)
    
    Dim objShape  As Shape
    For Each objShape In objSheet.Shapes
        Call objShape.Delete
    Next
    Call objSheet.Cells.Delete
    '�Ō�̃Z�����C������
    Call objSheet.Cells.Parent.UsedRange
End Sub

'*****************************************************************************
'[�T�v] Undo�{�^���̏����擾����
'[����] �Ȃ�
'[�ߒl] Undo�{�^����TooltipText
'*****************************************************************************
Public Function GetUndoStr() As String
    With CommandBars.FindControl(, 128) 'Undo�{�^��
        If .Enabled Then
            If .ListCount = 1 Then
                'Undo��1��ނ̎���Undo�R�}���h
                GetUndoStr = .TooltipText
            End If
        End If
    End With
End Function

