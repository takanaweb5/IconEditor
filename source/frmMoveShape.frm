VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMoveShape 
   Caption         =   "�}�`�̈ړ�"
   ClientHeight    =   2124
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   4812
   OleObjectBlob   =   "frmMoveShape.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmMoveShape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type TRect
    Top      As Double
    Height   As Double
    Left     As Double
    Width    As Double
End Type

Private Type TShapes  'Undo���
    Shapes() As TRect
End Type

Private udtShapes(1 To 100) As TShapes
Private lngUndoCount   As Long
Private objShapeRange  As ShapeRange
Private blnChange      As Boolean
Private blnOk          As Boolean
Private lngZoom        As Long
Private objDummy       As Shape

'*****************************************************************************
'[�T�v] �t�H�[�����[�h��
'*****************************************************************************
Private Sub UserForm_Initialize()
    Dim i  As Long
    
    '�Ăь��ɒʒm����
    blnFormLoad = True
    lngZoom = ActiveWindow.Zoom
    
    Set objShapeRange = Selection.ShapeRange
    
    '�u�O���b�h�ɂ��킹��v�̃`�F�b�N
    chkGrid.Value = CommandBars.GetPressedMso("SnapToGrid")
End Sub

'*****************************************************************************
'[�T�v] �t�H�[���A�����[�h��
'*****************************************************************************
Private Sub UserForm_Terminate()
    '�Ăь��ɒʒm����
    blnFormLoad = False
End Sub

'*****************************************************************************
'[�T�v] �~�{�^���Ńt�H�[������鎞�A�ύX�����ɖ߂�
'*****************************************************************************
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    Call objShapeRange.Select
    
    '�{�������ɖ߂�
    If ActiveWindow.Zoom <> lngZoom Then
        ActiveWindow.Zoom = lngZoom
    End If
    
    '�ύX���Ȃ���΃t�H�[�������
    If blnChange = False Then
        Exit Sub
    End If
        
    '�~�{�^���Ńt�H�[������鎞�A�t�H�[������Ȃ�
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Exit Sub
    End If
    
    If Not (objDummy Is Nothing) Then
        Call objDummy.Delete
    End If
    
    '�O���[�v�������}�`�̉���
    Call UnGroupSelection(objShapeRange).Select
    
    Call SetOnUndo("�}�`������")
End Sub

'*****************************************************************************
'[�T�v] �n�j�{�^��������
'*****************************************************************************
Private Sub cmdOK_Click()
    blnOk = True
    Call Unload(Me)
End Sub

'*****************************************************************************
'[�C�x���g]�@KeyDown
'[ �T  �v ]�@�J�[�\���L�[�ňړ����ύX������
'*****************************************************************************
Private Sub cmdOK_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call UserForm_KeyDown(KeyCode, Shift)
End Sub
Private Sub cmdCancel_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call UserForm_KeyDown(KeyCode, Shift)
End Sub
Private Sub chkGrid_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call UserForm_KeyDown(KeyCode, Shift)
End Sub
Private Sub cmdAlign_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call UserForm_KeyDown(KeyCode, Shift)
End Sub
Private Sub fraKeyCapture_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call UserForm_KeyDown(KeyCode, Shift)
End Sub

'*****************************************************************************
'[�T�v] �J�[�\���L�[�ňړ����ύX������
'*****************************************************************************
Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim blnGrid As Boolean
        
    '[Ctrl]+Z ����������Ă��鎞�AUndo���s��
    If (Shift = 2) And (KeyCode = vbKeyZ) Then
        Call PopUndoInfo
        Call fraKeyCapture.SetFocus
        KeyCode = 0
        Exit Sub
    End If
    
    Select Case (KeyCode)
    Case vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown, vbKeyPageUp, vbKeyPageDown, vbKeyHome
        Call fraKeyCapture.SetFocus
    Case Else
        Exit Sub
    End Select
    
    'Alt��������Ă���΃X�N���[��
    If GetKeyState(vbKeyMenu) < 0 Then
        Select Case (KeyCode)
        Case vbKeyLeft
            Call ActiveWindow.SmallScroll(, , , 1)
        Case vbKeyRight
            Call ActiveWindow.SmallScroll(, , 1)
        Case vbKeyUp
            Call ActiveWindow.SmallScroll(, 1)
        Case vbKeyDown
            Call ActiveWindow.SmallScroll(1)
        End Select
        Exit Sub
    End If
    
    'Zoom
    Select Case (KeyCode)
    Case vbKeyHome, vbKeyPageUp, vbKeyPageDown
        Call objShapeRange.Select
        Select Case (KeyCode)
        Case vbKeyHome
            If ActiveWindow.Zoom = lngZoom Then
                'Excel�̋@�\�𗘗p���āA�}�`��\���ł���ʒu�ɉ�ʂ��X�N���[�������邽��
                If lngZoom > 100 Then
                    ActiveWindow.Zoom = ActiveWindow.Zoom - 10
                Else
                    ActiveWindow.Zoom = ActiveWindow.Zoom + 10
                End If
            End If
            ActiveWindow.Zoom = lngZoom
        Case vbKeyPageUp
            ActiveWindow.Zoom = WorksheetFunction.min(ActiveWindow.Zoom + 10, 400)
        Case vbKeyPageDown
            ActiveWindow.Zoom = WorksheetFunction.max(ActiveWindow.Zoom - 10, 10)
        End Select
        If Not (objDummy Is Nothing) Then
            Call objDummy.Select
        End If
        Exit Sub
    End Select
    
    '�ύX�O�̏���ۑ�
    Call SaveBeforeChange
    
    '[Ctrl]Key����������Ă��� or �O���b�h�ɂ��킹�邪�`�F�b�N����Ă��Ȃ�
    If (GetKeyState(vbKeyControl) < 0) Or chkGrid.Value = False Then
        blnGrid = False
    Else
        blnGrid = True
    End If
    
On Error GoTo ErrHandle
    Dim dblSave    As Double
    Dim strMove    As String
    
    If GetKeyState(vbKeyShift) < 0 Then
        '�}�`�̑傫����ύX
        If GetKeyState(vbKeyZ) < 0 Then
            Select Case (KeyCode)
            Case vbKeyLeft
                Call ChangeShapesWidth(objShapeRange, 1, blnGrid, True)
                strMove = "Left"
            Case vbKeyRight
                Call ChangeShapesWidth(objShapeRange, -1, blnGrid, True)
            Case vbKeyUp
                Call ChangeShapesHeight(objShapeRange, 1, blnGrid, True)
                strMove = "Up"
            Case vbKeyDown
                Call ChangeShapesHeight(objShapeRange, -1, blnGrid, True)
            End Select
        Else
            Select Case (KeyCode)
            Case vbKeyLeft
                Call ChangeShapesWidth(objShapeRange, -1, blnGrid, False)
            Case vbKeyRight
                Call ChangeShapesWidth(objShapeRange, 1, blnGrid, False)
                strMove = "Right"
            Case vbKeyUp
                Call ChangeShapesHeight(objShapeRange, -1, blnGrid, False)
            Case vbKeyDown
                Call ChangeShapesHeight(objShapeRange, 1, blnGrid, False)
                strMove = "Down"
            End Select
        End If
    Else
        '�}�`���ړ�
        Select Case (KeyCode)
        Case vbKeyLeft
            Call MoveShapesLR(objShapeRange, -1, blnGrid)
            strMove = "Left"
        Case vbKeyRight
            Call MoveShapesLR(objShapeRange, 1, blnGrid)
            strMove = "Right"
        Case vbKeyUp
            Call MoveShapesUD(objShapeRange, -1, blnGrid)
            strMove = "Up"
        Case vbKeyDown
            Call MoveShapesUD(objShapeRange, 1, blnGrid)
            strMove = "Down"
        End Select
    End If

    '�I��̈悪��ʂ�����������ʂ��X�N���[��
    Dim i As Long
    If ActiveWindow.FreezePanes = False And ActiveWindow.Split = False Then '��ʕ����̂Ȃ���
        If strMove <> "" Then
            With GetShapeRangeRange(objShapeRange)
                Select Case (strMove)
                Case "Left"
                    i = WorksheetFunction.max(.Column - 1, 1)
                    If IntersectRange(ActiveWindow.VisibleRange, Columns(i)) Is Nothing Then
                        Call ActiveWindow.SmallScroll(, , , 1)
                    End If
                Case "Right"
                    i = WorksheetFunction.min(.Column + .Columns.Count, Columns.Count)
                    If IntersectRange(ActiveWindow.VisibleRange, Columns(i)) Is Nothing Then
                        Call ActiveWindow.SmallScroll(, , 1)
                    End If
                Case "Up"
                    i = WorksheetFunction.max(.Row - 1, 1)
                    If IntersectRange(ActiveWindow.VisibleRange, Rows(i)) Is Nothing Then
                        Call ActiveWindow.SmallScroll(, 1)
                    End If
                Case "Down"
                    i = WorksheetFunction.min(.Row + .Rows.Count, Rows.Count)
                    If IntersectRange(ActiveWindow.VisibleRange, Rows(i)) Is Nothing Then
                        Call ActiveWindow.SmallScroll(1)
                    End If
                End Select
            End With
            
            '�`��̎c������������
            Application.ScreenUpdating = True
        End If
    End If
ErrHandle:
End Sub

'*****************************************************************************
'[�T�v] ShapeRange�̎l���̃Z���͈͂��擾����
'[����] ShapeRange�I�u�W�F�N�g
'[�ߒl] �Z���͈�
'*****************************************************************************
Private Function GetShapeRangeRange(ByRef objShapeRange As ShapeRange) As Range
    Dim i As Long
    Set GetShapeRangeRange = GetNearlyRange(objShapeRange(1))
    
    For i = 2 To objShapeRange.Count
        Set GetShapeRangeRange = Range(GetShapeRangeRange, GetNearlyRange(objShapeRange(i)))
    Next
End Function

'*****************************************************************************
'[�T�v] �ύX�O�̏���ۑ�����
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub SaveBeforeChange()
    If blnChange = False Then
        '�A���h�D�p�Ɍ��̏�Ԃ�ۑ�����
'        Call SaveUndoInfo(E_ShapeSize, objShapeRange)
        Set objShapeRange = GroupSelection(objShapeRange)
        
        If Val(Application.Version) >= 12 Then
            Set objDummy = ActiveSheet.Shapes.AddLine(0, 0, 0, 0)
            Call objDummy.Select
        Else
            Call objShapeRange.Select
        End If

        blnChange = True
        '����{�^���𖳌��ɂ���
        Call EnableMenuItem(GetSystemMenu(FindWindow("ThunderDFrame", Me.Caption), False), SC_CLOSE, (MF_BYCOMMAND Or MF_GRAYED))
    End If

    Call PushUndoInfo
End Sub
    
'*****************************************************************************
'[�T�v] �ʒu����ۑ�����
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub PushUndoInfo()
    Dim i  As Long
    
    'Undo�ۑ����̍ő�𒴂�����
    If lngUndoCount = UBound(udtShapes) Then
        For i = 2 To UBound(udtShapes)
            udtShapes(i - 1) = udtShapes(i)
        Next
        lngUndoCount = lngUndoCount - 1
    End If
    
    lngUndoCount = lngUndoCount + 1
    With udtShapes(lngUndoCount)
        ReDim .Shapes(1 To objShapeRange.Count)
        For i = 1 To objShapeRange.Count
            .Shapes(i).Height = objShapeRange(i).Height
            .Shapes(i).Width = objShapeRange(i).Width
            .Shapes(i).Top = objShapeRange(i).Top
            .Shapes(i).Left = objShapeRange(i).Left
        Next
    End With
End Sub

'*****************************************************************************
'[�T�v] �ʒu���𕜌�����
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub PopUndoInfo()
    Dim i  As Long
    
    If lngUndoCount = 0 Then
        Exit Sub
    End If
    
    With udtShapes(lngUndoCount)
        For i = 1 To objShapeRange.Count
            objShapeRange(i).Height = .Shapes(i).Height
            objShapeRange(i).Width = .Shapes(i).Width
            objShapeRange(i).Top = .Shapes(i).Top
            objShapeRange(i).Left = .Shapes(i).Left
        Next
    End With

    lngUndoCount = lngUndoCount - 1
End Sub

'*****************************************************************************
'[�T�v] �}�`�����E�Ɉړ�����
'[����] objShapes:�}�`
'       lngSize:�ύX�T�C�Y(Pixel)
'       blnFitGrid:�g���ɂ��킹�邩
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub MoveShapesLR(ByRef objShapes As ShapeRange, ByVal lngSize As Long, ByVal blnFitGrid As Boolean)
    Dim objShape    As Shape   '�}�`
    Dim lngLeft      As Long
    Dim lngRight     As Long
    Dim lngNewLeft   As Long
    Dim lngNewRight  As Long
        
    '�g���ɂ��킹�邩
    If blnFitGrid = True Then
        '�}�`�̐��������[�v
        For Each objShape In objShapes
            lngLeft = Round(objShape.Left / DPIRatio)
            lngRight = Round((objShape.Left + objShape.Width) / DPIRatio)
            
            If lngSize < 0 Then
                lngNewLeft = GetLeftGrid(lngLeft, objShape.TopLeftCell.EntireColumn)
                If lngNewLeft < lngLeft Then
                   objShape.Left = lngNewLeft * DPIRatio
                End If
            Else
                lngNewRight = GetRightGrid(lngRight, objShape.BottomRightCell.EntireColumn)
                If lngNewRight > lngRight Then
                   objShape.Left = (lngLeft + lngNewRight - lngRight) * DPIRatio
                End If
            End If
        Next objShape
    Else
        '�s�N�Z���P�ʂ̈ړ����s��
        Call objShapes.IncrementLeft(lngSize * DPIRatio)
    End If
End Sub

'*****************************************************************************
'[�T�v] �}�`���㉺�Ɉړ�����
'[����] objShapes:�}�`
'       lngSize:�ύX�T�C�Y(Pixel)
'       blnFitGrid:�g���ɂ��킹�邩
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub MoveShapesUD(ByRef objShapes As ShapeRange, ByVal lngSize As Long, ByVal blnFitGrid As Boolean)
    Dim objShape     As Shape   '�}�`
    Dim lngTop       As Long
    Dim lngBottom    As Long
    Dim lngNewTop    As Long
    Dim lngNewBottom As Long
    
    '�g���ɂ��킹�邩
    If blnFitGrid = True Then
        '�}�`�̐��������[�v
        For Each objShape In objShapes
            lngTop = Round(objShape.Top / DPIRatio)
            lngBottom = Round((objShape.Top + objShape.Height) / DPIRatio)
            
            If lngSize < 0 Then
                lngNewTop = GetTopGrid(lngTop, objShape.TopLeftCell.EntireRow)
                If lngNewTop < lngTop Then
                   objShape.Top = lngNewTop * DPIRatio
                End If
            Else
                lngNewBottom = GetBottomGrid(lngBottom, objShape.BottomRightCell.EntireRow)
                If lngNewBottom > lngBottom Then
                   objShape.Top = (lngTop + lngNewBottom - lngBottom) * DPIRatio
                End If
            End If
        Next objShape
    Else
        '�s�N�Z���P�ʂ̈ړ����s��
        Call objShapes.IncrementTop(lngSize * DPIRatio)
    End If
End Sub

'*****************************************************************************
'[�T�v] �}�`�̃T�C�Y�ύX
'[����] objShapes:�}�`
'       lngSize:�ύX�T�C�Y(Pixel)
'       blnFitGrid:�g���ɂ��킹�邩
'       blnTopLeft:���܂��͏�����ɕω�������
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub ChangeShapesWidth(ByRef objShapes As ShapeRange, ByVal lngSize As Long, ByVal blnFitGrid As Boolean, Optional ByVal blnTopLeft As Boolean = False)
    Dim objShape     As Shape
    Dim lngLeft      As Long
    Dim lngRight     As Long
    Dim lngOldWidth  As Long
    Dim lngNewWidth  As Long
    Dim lngNewLeft   As Long
    Dim lngNewRight  As Long
    
    '�}�`�̐��������[�v
    For Each objShape In objShapes
        lngOldWidth = Round(objShape.Width / DPIRatio)
        lngLeft = Round(objShape.Left / DPIRatio)
        lngRight = Round((objShape.Left + objShape.Width) / DPIRatio)
        
        '�g���ɂ��킹�邩
        If blnFitGrid = True Then
            If blnTopLeft = True Then
                If lngSize > 0 Then
                    lngNewLeft = GetLeftGrid(lngLeft, objShape.TopLeftCell.EntireColumn)
                Else
                    lngNewLeft = GetRightGrid(lngLeft, objShape.TopLeftCell.EntireColumn)
                End If
                lngNewWidth = lngRight - lngNewLeft
            Else
                If lngSize < 0 Then
                    lngNewRight = GetLeftGrid(lngRight, objShape.BottomRightCell.EntireColumn)
                Else
                    lngNewRight = GetRightGrid(lngRight, objShape.BottomRightCell.EntireColumn)
                End If
                lngNewWidth = lngNewRight - lngLeft
            End If
            If lngNewWidth < 0 Then
                lngNewWidth = 0
            End If
        Else
            '�s�N�Z���P�ʂ̕ύX������
            If lngOldWidth + lngSize >= 0 Then
                If blnTopLeft = True And lngLeft = 0 And lngSize > 0 Then
                    lngNewWidth = lngOldWidth
                Else
                    lngNewWidth = lngOldWidth + lngSize
                End If
            Else
                lngNewWidth = lngOldWidth
            End If
        End If
    
        If lngSize > 0 And blnTopLeft = True Then
            objShape.Left = (lngRight - lngNewWidth) * DPIRatio
        End If
        objShape.Width = lngNewWidth * DPIRatio
        
        'Excel2007�̃o�O�Ή�
        If Round(objShape.Width / DPIRatio) <> lngNewWidth Then
            objShape.Width = (lngNewWidth + lngSize) * DPIRatio
        End If
        
        If Round(objShape.Width / DPIRatio) <> lngOldWidth Then
            If blnTopLeft = True Then
                objShape.Left = (lngRight - lngNewWidth) * DPIRatio
            Else
                objShape.Left = lngLeft * DPIRatio
            End If
        End If
    Next objShape
End Sub

'*****************************************************************************
'[�T�v] �}�`�̃T�C�Y�ύX
'[����] objShapes:�}�`
'       lngSize:�ύX�T�C�Y(Pixel)
'       blnFitGrid:�g���ɂ��킹�邩
'       blnTopLeft:���܂��͏�����ɕω�������
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub ChangeShapesHeight(ByRef objShapes As ShapeRange, ByVal lngSize As Long, ByVal blnFitGrid As Boolean, Optional ByVal blnTopLeft As Boolean = False)
    Dim objShape     As Shape
    Dim lngTop       As Long
    Dim lngBottom    As Long
    Dim lngOldHeight As Long
    Dim lngNewHeight As Long
    Dim lngNewTop    As Long
    Dim lngNewBottom As Long
    
    '�}�`�̐��������[�v
    For Each objShape In objShapes
        lngOldHeight = Round(objShape.Height / DPIRatio)
        lngTop = Round(objShape.Top / DPIRatio)
        lngBottom = Round((objShape.Top + objShape.Height) / DPIRatio)
        
        '�g���ɂ��킹�邩
        If blnFitGrid = True Then
            If blnTopLeft = True Then
                If lngSize > 0 Then
                    lngNewTop = GetTopGrid(lngTop, objShape.TopLeftCell.EntireRow)
                Else
                    lngNewTop = GetBottomGrid(lngTop, objShape.TopLeftCell.EntireRow)
                End If
                lngNewHeight = lngBottom - lngNewTop
            Else
                If lngSize < 0 Then
                    lngNewBottom = GetTopGrid(lngBottom, objShape.BottomRightCell.EntireRow)
                Else
                    lngNewBottom = GetBottomGrid(lngBottom, objShape.BottomRightCell.EntireRow)
                End If
                lngNewHeight = lngNewBottom - lngTop
            End If
            If lngNewHeight < 0 Then
                lngNewHeight = 0
            End If
        Else
            '�s�N�Z���P�ʂ̕ύX������
            If lngOldHeight + lngSize >= 0 Then
                If blnTopLeft = True And lngTop = 0 And lngSize > 0 Then
                    lngNewHeight = lngOldHeight
                Else
                    lngNewHeight = lngOldHeight + lngSize
                End If
            Else
                lngNewHeight = lngOldHeight
            End If
        End If
    
        If lngSize > 0 And blnTopLeft = True Then
            objShape.Top = (lngBottom - lngNewHeight) * DPIRatio
        End If
        objShape.Height = lngNewHeight * DPIRatio
        
        'Excel2007�̃o�O�Ή�
        If Round(objShape.Height / DPIRatio) <> lngNewHeight Then
            objShape.Height = (lngNewHeight + lngSize) * DPIRatio
        End If
        
        If Round(objShape.Height / DPIRatio) <> lngOldHeight Then
            If blnTopLeft = True Then
                objShape.Top = (lngBottom - lngNewHeight) * DPIRatio
            Else
                objShape.Top = lngTop * DPIRatio
            End If
        End If
    Next objShape
End Sub

