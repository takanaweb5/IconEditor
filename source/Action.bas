Attribute VB_Name = "Action"
Option Explicit

Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
Private Declare PtrSafe Function GetKeyState Lib "user32" (ByVal lngVirtKey As Long) As Integer
Public FFormLoad As Boolean

'*****************************************************************************
'[�T�v] ImageMso���擾
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub ImageMso�擾()
    If CheckSelection <> E_Range Then
        Call MsgBox("�摜��Ǎ��ރZ����I�����Ă�������", vbCritical)
        Exit Sub
    End If
    
    Dim strImgMso As String
    strImgMso = InputBox("ImageMso����͂��Ă�������" & vbCrLf & vbCrLf & "�� Copy")
    If strImgMso = "" Then Exit Sub
    
    '�`�F�b�N
    On Error Resume Next
    Call CommandBars.GetImageMso(strImgMso, 32, 32)
    If Err.Number <> 0 Then
        Call MsgBox("ImageMso������Ă��܂�")
        Exit Sub
    End If
    
    Dim WidthAndHeight As Variant
    Dim strInput As String
    Dim lngWidth As Long
    Dim lngHeight As Long
    Do While True
        strInput = InputBox("��,��������͂��Ă�������" & vbCrLf & vbCrLf & "�� 32,32", , "32,32")
        If strInput = "" Then
            Exit Sub
        End If
        WidthAndHeight = Split(strInput, ",")
        If UBound(WidthAndHeight) = 1 Then
            If IsNumeric(WidthAndHeight(0)) And IsNumeric(WidthAndHeight(1)) Then
                lngWidth = WidthAndHeight(0)
                lngHeight = WidthAndHeight(1)
                Exit Do
            End If
        End If
    Loop
    
On Error GoTo ErrHandle
    Dim img As New CImage
    Call img.GetPixelsFromHBITMAP(CommandBars.GetImageMso(strImgMso, lngWidth, lngHeight).Handle)
    Application.ScreenUpdating = False
    Call SaveUndoInfo(ActiveCell.Resize(img.Height, img.Width))
    Call img.SetPixelsToRange(ActiveCell)
    Call ActiveCell.Resize(img.Height, img.Width).Select
    Call SetOnUndo("ImageMso�擾")
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[�T�v] �摜��Ǎ���
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub �摜�Ǎ�()
On Error GoTo ErrHandle
    If CheckSelection <> E_Range Then
        Call MsgBox("�摜��Ǎ��ރZ����I�����Ă�������")
        Exit Sub
    End If
    Dim vDBName As Variant
    vDBName = Application.GetOpenFilename("PNG,*.png,�A�C�R��,*.ico,�r�b�g�}�b�v,*.bmp,�S�Ẵt�@�C��,*.*")
    If vDBName = False Then
        Exit Sub
    End If
    
    Dim img As New CImage
    Call img.LoadImageFromFile(vDBName)
    Application.ScreenUpdating = False
    Call SaveUndoInfo(ActiveCell.Resize(img.Height, img.Width))
    Call img.SetPixelsToRange(ActiveCell)
    Call ActiveCell.Resize(img.Height, img.Width).Select
    Call SetOnUndo("�摜�Ǎ�")
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[�T�v] �摜��ۑ�
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub �摜�ۑ�()
On Error GoTo ErrHandle
    If CheckSelection <> E_Range Then
        Call MsgBox("�摜��I�����Ă�������")
        Exit Sub
    End If
    If Selection.Rows.Count = 1 And Selection.Columns.Count = 1 Then
        Call MsgBox("�摜��I�����Ă�������")
        Exit Sub
    End If
    If Selection.Rows.Count = Rows.Count Or Selection.Columns.Count = Columns.Count Then
        Call MsgBox("���ׂĂ̍s�܂��͗�̑I�����͎��s�o���܂���")
        Exit Sub
    End If
    
    Dim vDBName As Variant
    vDBName = Application.GetSaveAsFilename("", "PNG,*.png,�A�C�R��,*.ico,�r�b�g�}�b�v,*.bmp,�S�Ẵt�@�C��,*.*")
    If vDBName = False Then Exit Sub
    
    Dim img As New CImage
    Call img.GetPixelsFromRange(Selection)
    Call img.SaveImageToFile(vDBName)
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[�T�v] �N���b�v�{�[�h�̉摜��ۑ�
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub Clipbord�摜�ۑ�()
On Error GoTo ErrHandle
    If Not ClipboardHasBitmap() Then
        Call MsgBox("�N���b�v�{�[�h�ɉ摜������܂���")
        Exit Sub
    End If
    
    Dim vDBName As Variant
    vDBName = Application.GetSaveAsFilename("", "PNG,*.png,�A�C�R��,*.ico,�r�b�g�}�b�v,*.bmp,�S�Ẵt�@�C��,*.*")
    If vDBName = False Then Exit Sub
    
    Dim img As New CImage
    Call img.LoadBMPFromClipbord
    Call img.SaveImageToFile(vDBName)
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[�T�v] �㉺���]
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub �㉺���]()
On Error GoTo ErrHandle
    If CheckSelection <> E_Range Then Exit Sub
    If Selection.Rows.Count = Rows.Count Or Selection.Columns.Count = Columns.Count Then
        Call MsgBox("���ׂĂ̍s�܂��͗�̑I�����͎��s�o���܂���")
        Exit Sub
    End If
    
    Dim img As New CImage
    Call img.GetPixelsFromRange(Selection)
    Call img.FlipHorizontal
    Application.ScreenUpdating = False
    Call SaveUndoInfo(Selection)
    Call img.SetPixelsToRange(Selection)
    Call SetOnUndo("���]")
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[�T�v] ���E���]
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub ���E���]()
On Error GoTo ErrHandle
    If CheckSelection <> E_Range Then Exit Sub
    If Selection.Rows.Count = Rows.Count Or Selection.Columns.Count = Columns.Count Then
        Call MsgBox("���ׂĂ̍s�܂��͗�̑I�����͎��s�o���܂���")
        Exit Sub
    End If
    
    Dim img As New CImage
    Call img.GetPixelsFromRange(Selection)
    Call img.FlipVertical
    Application.ScreenUpdating = False
    Call SaveUndoInfo(Selection)
    Call img.SetPixelsToRange(Selection)
    Call SetOnUndo("���]")
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[�T�v] 90�x��]
'[����] Mode 1:�I��͈͂���],2:�N���b�v�{�[�h�̗̈����]
'       Angle:90 or -90
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub ��](ByVal lngMode As Long, ByVal lngAngle As Long)
On Error GoTo ErrHandle
    If CheckSelection <> E_Range Then Exit Sub
    If Selection.Rows.Count = Rows.Count Or Selection.Columns.Count = Columns.Count Then
        Call MsgBox("���ׂĂ̍s�܂��͗�̑I�����͎��s�o���܂���")
        Exit Sub
    End If
    
    Dim objCopyRange As Range
    If lngMode = 1 Then
        If Selection.Rows.Count <> Selection.Columns.Count Then
            Call MsgBox("���ƍ������s��v�̂��ߎ��s�ł��܂���" & vbCrLf & "�\�t�R�}���h�̒��̉�]�����s���Ă�������")
            Exit Sub
        Else
            Set objCopyRange = Selection
        End If
    Else
        Set objCopyRange = GetCopyRange()
        If objCopyRange Is Nothing Then
            Call MsgBox("��]������̈���R�s�[���Ă�����s���Ă�������")
            Exit Sub
        End If
    End If

    Dim img As New CImage
    Call img.GetPixelsFromRange(objCopyRange)
    Call img.Rotate(lngAngle)
    Application.ScreenUpdating = False
    Call SaveUndoInfo(ActiveCell.Resize(img.Height, img.Width))
    Call img.SetPixelsToRange(Selection)
    Call ActiveCell.Resize(img.Height, img.Width).Select
    Call SetOnUndo("��]")
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[�T�v] �N���b�v�{�[�h�̉摜��Ǎ���
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub Clipbord�摜�Ǎ�()
On Error GoTo ErrHandle
    If Not ClipboardHasBitmap() Then
        Call MsgBox("�N���b�v�{�[�h�ɉ摜������܂���")
        Exit Sub
    End If
    If CheckSelection <> E_Range Then
        Call MsgBox("�摜��Ǎ��ރZ����I�����Ă�������")
        Exit Sub
    End If
    
    Dim img As New CImage
    Call img.LoadBMPFromClipbord
    Dim strDefault As String
    strDefault = img.Width & "," & img.Height
    
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim WidthAndHeight As Variant
    Dim strInput As String
    Dim objSelection
    Set objSelection = Selection
    Do While True
        strInput = InputBox("��,��������͂��Ă�������", , strDefault)
        If strInput = "" Then
            Exit Sub
        End If
        WidthAndHeight = Split(strInput, ",")
        If UBound(WidthAndHeight) = 1 Then
            If IsNumeric(WidthAndHeight(0)) And IsNumeric(WidthAndHeight(1)) Then
                lngWidth = WidthAndHeight(0)
                lngHeight = WidthAndHeight(1)
                Exit Do
            End If
        End If
    Loop
    
    Call img.Resize(lngWidth, lngHeight)
    Application.ScreenUpdating = False
    Call SaveUndoInfo(ActiveCell.Resize(img.Height, img.Width))
    Call img.SetPixelsToRange(Selection)
    Call ActiveCell.Resize(img.Height, img.Width).Select
    Call SetOnUndo("�摜�Ǎ�")
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[�T�v] �摜�ɕϊ�
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub �摜�ɕϊ�()
On Error GoTo ErrHandle
    If CheckSelection <> E_Range Then Exit Sub
    If Selection.Rows.Count = Rows.Count Or Selection.Columns.Count = Columns.Count Then
        Call MsgBox("���ׂĂ̍s�܂��͗�̑I�����͎��s�o���܂���")
        Exit Sub
    End If
    
    Dim img As New CImage
    Call img.GetPixelsFromRange(Selection)
    Call img.SaveImageToClipbord
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[�T�v] Shape���摜�ɕϊ����ēǍ���
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub Shape�Ǎ�()
On Error GoTo ErrHandle
    If CheckSelection <> E_Shape Then
        Call MsgBox("�I�[�g�V�F�C�v���I������Ă��܂���")
        Exit Sub
    End If
    Dim objSelection
    Set objSelection = Selection
    
    Dim objCell As Range
    Call ActiveCell.Select
    Set objCell = SelectCell("�摜��Ǎ��ރZ����I�����Ă�������", ActiveCell)
    If objCell Is Nothing Then
        Exit Sub
    End If
    
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim WidthAndHeight As Variant
    Dim strInput As String
    Do While True
        strInput = InputBox("��,������" & MAX_WIDTH & "Pixel�����œ��͂��Ă�������")
        If strInput = "" Then
            Exit Do
        End If
        WidthAndHeight = Split(strInput, ",")
        If UBound(WidthAndHeight) = 1 Then
            If IsNumeric(WidthAndHeight(0)) And IsNumeric(WidthAndHeight(1)) Then
                lngWidth = WidthAndHeight(0)
                lngHeight = WidthAndHeight(1)
                If lngWidth <= MAX_WIDTH And lngHeight <= MAX_HEIGHT Then
                    Exit Do
                End If
            End If
        End If
    Loop
    Dim objWkShape As Shape
    Set objWkShape = GroupShape(objSelection.ShapeRange(1))
    
    '72(Excel�̃f�t�H���g��DPI),96(Windows�摜�̃f�t�H���g��DPI)
    objWkShape.Width = (lngWidth - 1) * 72 / 96
    objWkShape.Height = (lngHeight - 1) * 72 / 96
    Call objWkShape.Copy
    
    Dim img As New CImage
    Call img.LoadBMPFromClipbord
    Call img.Resize(lngWidth, lngHeight)
    Application.ScreenUpdating = False
    Call SaveUndoInfo(objCell.Resize(img.Height, img.Width))
    Call img.SetPixelsToRange(objCell)
    Call objCell.Resize(img.Height, img.Width).Select
    If Not (objWkShape Is Nothing) Then
        Call objWkShape.Delete
    End If
    Call SetOnUndo("�摜�Ǎ�")
Exit Sub
ErrHandle:
    If Not (objWkShape Is Nothing) Then
        Call objWkShape.Delete
    End If
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[�T�v] ��]���Ă���Shape�͕��ƍ������ςɂȂ�̂�Group�����Ă܂Ƃ��ɂȂ�悤�ɂ���
'[����] �O���[�v���O��Shape
'[�ߒl] �O���[�v�����Shape
'*****************************************************************************
Private Function GroupShape(ByRef objShape As Shape) As Shape
    ReDim lngIDArray(1 To 2) As Variant
    '�N���[�����Q�쐬���O���[�v������
    With objShape.Duplicate
        .Top = objShape.Top
        .Left = objShape.Left
        lngIDArray(1) = .ID
    End With
    With objShape.Duplicate
        .Top = objShape.Top
        .Left = objShape.Left
        '�����ɂ���
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        lngIDArray(2) = .ID
    End With
    Set GroupShape = GetShapeRangeFromID(lngIDArray).Group
End Function

'*****************************************************************************
'[ �֐��� ]�@GetShapeRangeFromID
'[ �T  �v ]�@Shpes�I�u�W�F�N�g��ID����ShapeRange�I�u�W�F�N�g���擾
'[ ��  �� ]�@ID�̔z��
'[ �߂�l ]�@ShapeRange�I�u�W�F�N�g
'*****************************************************************************
Private Function GetShapeRangeFromID(ByRef lngID As Variant) As ShapeRange
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
'[�T�v] �����F�̋���
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub �����F����()
On Error GoTo ErrHandle
    If CheckSelection <> E_Range Then Exit Sub
    If Selection.Rows.Count = Rows.Count Or Selection.Columns.Count = Columns.Count Then
        Call MsgBox("���ׂĂ̍s�܂��͗�̑I�����͎��s�o���܂���")
        Exit Sub
    End If
    
    '�I��͈͂̏d����r��
    Dim objSelection As Range
    Set objSelection = ReSelectRange(Selection)
    
    Application.ScreenUpdating = False
    Call SaveUndoInfo(Selection)
    Dim objCell As Range
    For Each objCell In objSelection
        Call ColorToCell(objCell, CellToColor(objCell))
    Next
    Call SetOnUndo("�����F����")
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[�T�v] �F�̒u��
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub �F�̒u��()
On Error GoTo ErrHandle
    Dim objSelection As Range
    If CheckSelection = E_Range Then
        Set objSelection = Selection
    Else
        Set objSelection = ActiveCell
    End If
    
    Dim objCanvas As Range
    Set objCanvas = SelectCell("�L�����o�X�͈̔͂�I�����Ă�������", objSelection)
    If objCanvas Is Nothing Then
        Exit Sub
    Else
        If objCanvas.Rows.Count = Rows.Count Or objCanvas.Columns.Count = Columns.Count Then
            Call MsgBox("���ׂĂ̍s�܂��͗�̑I�����͎��s�o���܂���")
            Exit Sub
        End If
    End If
    '�I��͈͂̏d����r��
    Set objCanvas = ReSelectRange(objCanvas)
    
    Dim objCell As Range
    Dim SrcColor As TRGBQuad
    Set objCell = SelectCell("�u���O�̐F�̃Z����I�����Ă�������", ActiveCell)
    If objCell Is Nothing Then
        Exit Sub
    End If
    SrcColor = CellToColor(objCell(1))
    
    Dim DstColor As TRGBQuad
    Set objCell = SelectCell("�u����̐F�̃Z����I�����Ă�������", ActiveCell)
    If objCell Is Nothing Then
        Exit Sub
    End If
    DstColor = CellToColor(objCell(1))
    
    Application.ScreenUpdating = False
    Call SaveUndoInfo(Selection)
    For Each objCell In objCanvas
        If SameColor(SrcColor, CellToColor(objCell)) Then
            Call ColorToCell(objCell, DstColor)
        End If
    Next
    Call Selection.Select
    Call SetOnUndo("�F�̒u��")
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[�T�v] �����i�܂��͑��Ⴗ��j�F�̃Z����I��
'[����] True:�����F�̃Z����I���AFalse:�Ⴄ�F�̃Z����I��
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub ���F�I��(ByVal blnSameColor As Boolean)
On Error GoTo ErrHandle
    Dim objSelection As Range
    Dim objCell As Range
    If CheckSelection = E_Range Then
        Set objSelection = Selection
    Else
        Set objSelection = ActiveCell
    End If
    Dim objCanvas As Range
    Set objCanvas = SelectCell("�L�����o�X�͈̔͂�I�����Ă�������", objSelection)
    If objCanvas Is Nothing Then
        Exit Sub
    Else
        If objCanvas.Rows.Count = Rows.Count Or objCanvas.Columns.Count = Columns.Count Then
            Call MsgBox("���ׂĂ̍s�܂��͗�̑I�����͎��s�o���܂���")
            Exit Sub
        End If
    End If
    
    Dim SelectColor  As TRGBQuad
    Dim strMsg As String
    If blnSameColor Then
        strMsg = "�I���������F�̃Z����I�����Ă�������"
    Else
        strMsg = "�I���������Ȃ��F�̃Z����I�����Ă�������"
    End If
    Set objCell = SelectCell(strMsg, ActiveCell)
    If objCell Is Nothing Then
        Exit Sub
    End If
    SelectColor = CellToColor(objCell(1))
    
    Dim objNewSelection As Range
    For Each objCell In objCanvas
        If SameColor(SelectColor, CellToColor(objCell)) = blnSameColor Then
            If objNewSelection Is Nothing Then
                Set objNewSelection = objCell
            Else
                Set objNewSelection = Application.Union(objNewSelection, objCell)
            End If
        End If
    Next
    If objNewSelection Is Nothing Then
        Call MsgBox("�Y���̃Z���͂���܂���ł���")
    Else
        Call objNewSelection.Select
    End If
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[�T�v] �I���Z���̔��]�Ȃ�
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub �I�𔽓]��()
    Static strLastSheet   As String '�O��̗̈�̕����p
    Static strLastAddress As String '�O��̗̈�̕����p
On Error GoTo ErrHandle
    Dim objSelection As Range
    Dim objUnSelect  As Range
    Dim objRange As Range
    Dim enmUnselectMode As EUnselectMode
    If CheckSelection() <> E_Range Then
        Exit Sub
    End If
    Set objSelection = Selection
    
    '����̈��I��������
    With frmUnSelect
        '�O��̕����p
        Call .SetLastSelect(strLastSheet, strLastAddress)
        '�t�H�[����\��
        Call .Show
        '�L�����Z����
        If FFormLoad = False Then
            Exit Sub
        End If
        enmUnselectMode = .Mode
        Select Case (enmUnselectMode)
        Case E_Unselect, E_Reverse, E_Intersect, E_Union
            Set objUnSelect = .SelectRange
        End Select
        Call Unload(frmUnSelect)
    End With

    Select Case (enmUnselectMode)
    Case E_Unselect  '�����
        Set objRange = MinusRange(objSelection, objUnSelect)
    Case E_Reverse   '���]
        Set objRange = UnionRange(MinusRange(objSelection, objUnSelect), MinusRange(objUnSelect, objSelection))
    Case E_Intersect '�i�荞��
        Set objRange = IntersectRange(objSelection, objUnSelect)
    Case E_Union     '�ǉ�
        Set objRange = UnionRange(objSelection, objUnSelect)
    End Select
    
    If Not (objRange Is Nothing) Then
        Call objRange.Select
    End If
ErrHandle:
    strLastSheet = ActiveSheet.Name
    strLastAddress = Selection.Address(False, False)
    If FFormLoad Then
        Call Unload(frmUnSelect)
    End If
End Sub

'*****************************************************************************
'[�T�v] �F�̓\�t��
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub �\�t��()
On Error GoTo ErrHandle
    If CheckSelection <> E_Range Then Exit Sub
    If Selection.Rows.Count = Rows.Count Or Selection.Columns.Count = Columns.Count Then
        Call MsgBox("���ׂĂ̍s�܂��͗�̑I�����͎��s�o���܂���")
        Exit Sub
    End If
    Dim objCopyRange  As Range
    Set objCopyRange = GetCopyRange()
    If objCopyRange Is Nothing Then Exit Sub
    If objCopyRange.Rows.Count = Rows.Count Or objCopyRange.Columns.Count = Columns.Count Then
        Call MsgBox("���ׂĂ̍s�܂��͗�̑I�����͎��s�o���܂���")
        Exit Sub
    End If
    
    Dim objSelection As Range
    Set objSelection = Selection
    Dim objColorCell As Range
    Dim ColorFlg As Long
    Dim lngMode As Long
    If objCopyRange.Count > 1 Then
        If objSelection.Areas.Count > 1 Then
            Call MsgBox("���̃R�}���h�͕����̑I��͈͂ɑ΂��Ď��s�ł��܂���")
            Exit Sub
        End If
        If FChecked(1) Then
            Set objColorCell = SelectCell("�ΏېF�̃Z����I�����Ă�������", ActiveCell)
            If objColorCell Is Nothing Then Exit Sub
            lngMode = 1
        End If
        If FChecked(2) Then
            Set objColorCell = SelectCell("���O�Ώۂ̐F�̃Z����I�����Ă�������", ActiveCell)
            If objColorCell Is Nothing Then Exit Sub
            lngMode = 2
        End If
    End If
    
    Application.ScreenUpdating = False
    If objCopyRange.Count = 1 Then
        Dim DstColor As TRGBQuad
        DstColor = CellToColor(objCopyRange)
        Call SaveUndoInfo(objSelection)
        Dim objCell As Range
        For Each objCell In objSelection
            Call ColorToCell(objCell, DstColor)
        Next
    Else
        Dim img As New CImage
        Call img.GetPixelsFromRange(objCopyRange)
        Call SaveUndoInfo(objSelection.Resize(img.Height, img.Width))
        Call img.SetPixelsToRange(objSelection, lngMode, objColorCell, FChecked(3))
        Call objSelection.Resize(img.Height, img.Width).Select
    End If
    Call SetOnUndo("�\�t��")
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[�T�v] �N���b�v�{�[�h�ɉ摜��ݒ肷��
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub Clipbord�摜�ݒ�()
On Error GoTo ErrHandle
    Dim img As New CImage
    Call img.GetPixelsFromRange(Selection)
    Call img.SaveImageToClipbord
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub


'*****************************************************************************
'[�T�v] �F��ARGB�𑝌�������
'[����] 1:�����A-1:����
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub �F����(ByVal lngUp As Long)
On Error GoTo ErrHandle
    If CheckSelection <> E_Range Then Exit Sub
    If FChecked(3) Or FChecked(4) Or FChecked(5) Or FChecked(6) Or FChecked(7) Then
    Else
        Call MsgBox("RGB����уA���t�@�l�̂�������`�F�b�N����Ă��܂���")
        Exit Sub
    End If
    If Selection.Rows.Count = Rows.Count Or Selection.Columns.Count = Columns.Count Then
        Call MsgBox("���ׂĂ̍s�܂��͗�̑I�����͎��s�o���܂���")
        Exit Sub
    End If

    If GetKeyState(vbKeyControl) < 0 Then
        lngUp = lngUp * 1
    Else
        lngUp = lngUp * 16
    End If
    
    '�I��͈͂̏d����r��
    Dim objSelection As Range
    Set objSelection = ReSelectRange(Selection)
    
    Dim objCell As Range
    Dim ARGB As TRGBQuad
    Application.ScreenUpdating = False
    Call SaveUndoInfo(Selection, "�F����")
    For Each objCell In objSelection
        ARGB = AdjustColor(CellToColor(objCell), lngUp, FChecked(4), FChecked(5), FChecked(6), FChecked(7))
        Call ColorToCell(objCell, ARGB)
    Next
    Call Selection.Select
    Call SetOnUndo("�F����")
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[�T�v] �摜��IPicture�ɕϊ�����
'[����] �Ȃ�
'[�ߒl] IPicture
'*****************************************************************************
Public Function Get�T���v���摜() As IPicture
On Error GoTo ErrHandle
    Dim img As New CImage
    Call img.GetPixelsFromRange(Selection)
    Set Get�T���v���摜 = img.SetToIPicture
Exit Function
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Function

'*****************************************************************************
'[�T�v] �h�ׂ�
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub �h�ׂ�()
On Error GoTo ErrHandle
    Dim objSelection As Range
    If CheckSelection = E_Range Then
        Set objSelection = Selection.Areas(1)
    Else
        Set objSelection = ActiveCell
    End If
    Dim objCanvas As Range
    Set objCanvas = SelectCell("�L�����o�X�͈̔͂�I�����Ă�������", objSelection)
    If objCanvas Is Nothing Then
        Exit Sub
    Else
        If objCanvas.Rows.Count = Rows.Count Or objCanvas.Columns.Count = Columns.Count Then
            Call MsgBox("���ׂĂ̍s�܂��͗�̑I�����͎��s�o���܂���")
            Exit Sub
        End If
    End If
    
    Dim objColorCell As Range
    Set objColorCell = SelectCell("�h�ׂ��F�̃Z����I�����Ă�������", ActiveCell)
    If objColorCell Is Nothing Then
        Exit Sub
    End If
    
    Dim objStartCell As Range
    Set objStartCell = SelectCell("�h�ׂ��J�n�Z����I�����Ă�������", ActiveCell)
    If objStartCell Is Nothing Then
        Exit Sub
    End If
    
    If Intersect(objCanvas, objStartCell) Is Nothing Then
        Call MsgBox("�h�ׂ��J�n�Z�����A�L�����o�X�̓����ɂ���܂���")
        Exit Sub
    End If
    
    Dim img As New CImage
    Call img.GetPixelsFromRange(objCanvas)
    Call img.Fill(objStartCell.Column - objCanvas.Column + 1, _
                  objStartCell.Row - objCanvas.Row + 1, _
                  objColorCell)
    Application.ScreenUpdating = False
    Call SaveUndoInfo(objCanvas)
    Call img.SetPixelsToRange(objCanvas)
    Call SetOnUndo("�h�ׂ�")
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

