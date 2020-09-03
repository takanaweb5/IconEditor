Attribute VB_Name = "Action"
Option Explicit
Option Private Module

Public Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
Private Declare PtrSafe Function GetKeyState Lib "user32" (ByVal lngVirtKey As Long) As Integer
Public FFormLoad As Boolean

'*****************************************************************************
'[�T�v] �N���A
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub �N���A()
On Error GoTo ErrHandle
    If CheckSelection <> E_Range Then Exit Sub
    Application.ScreenUpdating = False
    Call SaveUndoInfo(Selection)
    Call ClearRange(Selection)
'    Call ColorToCell(Selection, OleColorToARGB(0, 0), True)
    Call SetOnUndo("�N���A")
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

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
    Dim objSelection As Range
    Set objSelection = Selection
    
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
    
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim DestRange As Range
    If 1 < objSelection.Columns.Count And objSelection.Columns.Count <= 64 And _
       1 < objSelection.Rows.Count And objSelection.Rows.Count <= 64 Then
        lngWidth = objSelection.Columns.Count
        lngHeight = objSelection.Rows.Count
        Set DestRange = objSelection
    Else
        Dim WidthAndHeight As Variant
        Dim strInput As String
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
        Set DestRange = ActiveCell.Resize(lngHeight, lngWidth)
    End If
    If 16 <= lngWidth And lngWidth <= 64 And _
       16 <= lngHeight And lngHeight <= 64 Then
    Else
        Call MsgBox("������э����� 16�`64 �Ŏw�肵�Ă�������")
        Exit Sub
    End If

On Error GoTo ErrHandle
    Dim img As New CImage
    Call img.GetPixelsFromHBITMAP(CommandBars.GetImageMso(strImgMso, lngWidth, lngHeight).Handle)
    Application.ScreenUpdating = False
    Call SaveUndoInfo(DestRange)
    Call img.SetPixelsToRange(DestRange)
    Call DestRange.Select
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
    If Selection.Rows.Count = Rows.Count Or Selection.Columns.Count = Columns.Count Then
        Call MsgBox("���ׂĂ̍s�܂��͗�̑I�����͎��s�o���܂���")
        Exit Sub
    End If
    Dim objSelection As Range
    Set objSelection = Selection
    
    Dim vDBName As Variant
    vDBName = Application.GetOpenFilename("PNG,*.png,�A�C�R��,*.ico,�r�b�g�}�b�v,*.bmp,�S�Ẵt�@�C��,*.*")
    If vDBName = False Then
        Exit Sub
    End If
    
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim DestRange As Range
    If 1 < objSelection.Columns.Count And _
       1 < objSelection.Rows.Count Then
        lngWidth = objSelection.Columns.Count
        lngHeight = objSelection.Rows.Count
        Set DestRange = objSelection
    End If
    
    Dim img As New CImage
    Call img.LoadImageFromFile(vDBName, lngWidth, lngHeight)
    If DestRange Is Nothing Then
        Set DestRange = ActiveCell.Resize(img.Height, img.Width)
    End If
    
    Application.ScreenUpdating = False
    Call SaveUndoInfo(DestRange)
    Call img.SetPixelsToRange(DestRange)
    Call DestRange.Select
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
    If Selection.Rows.Count = Rows.Count Or Selection.Columns.Count = Columns.Count Then
        Call MsgBox("���ׂĂ̍s�܂��͗�̑I�����͎��s�o���܂���")
        Exit Sub
    End If
    Dim objSelection As Range
    Set objSelection = Selection
    
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim DestRange As Range
    If 1 < objSelection.Columns.Count And _
       1 < objSelection.Rows.Count Then
        lngWidth = objSelection.Columns.Count
        lngHeight = objSelection.Rows.Count
        Set DestRange = objSelection
    End If
    
    Dim img As New CImage
    Call img.LoadBMPFromClipbord(lngWidth, lngHeight)
    If DestRange Is Nothing Then
        Set DestRange = ActiveCell.Resize(img.Height, img.Width)
    End If
    
    Application.ScreenUpdating = False
    Call SaveUndoInfo(DestRange)
    Call img.SetPixelsToRange(DestRange)
    Call DestRange.Select
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
    
    Dim DestRange As Range
    Call ActiveCell.Select
    Set DestRange = SelectCell("�摜��Ǎ��ރZ����I�����Ă�������", ActiveCell)
    If DestRange Is Nothing Then
        Exit Sub
    End If
    
    Dim lngWidth As Long
    Dim lngHeight As Long
    If 1 < DestRange.Columns.Count And DestRange.Columns.Count <= 64 And _
       1 < DestRange.Rows.Count And DestRange.Rows.Count <= 64 Then
        lngWidth = DestRange.Columns.Count
        lngHeight = DestRange.Rows.Count
    Else
        Dim WidthAndHeight As Variant
        Dim strInput As String
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
        Set DestRange = DestRange.Resize(lngHeight, lngWidth)
    End If
    
    Dim objWkShape As Shape
    Set objWkShape = GroupShape(objSelection.ShapeRange(1))
    
    '72(Excel�̃f�t�H���g��DPI),96(Windows�摜�̃f�t�H���g��DPI)
    objWkShape.Width = (lngWidth - 1) * 72 / 96
    objWkShape.Height = (lngHeight - 1) * 72 / 96
    Call objWkShape.Copy
    
    Dim img As New CImage
    Call img.LoadBMPFromClipbord(lngWidth, lngHeight)
    Application.ScreenUpdating = False
    Call SaveUndoInfo(DestRange)
    Call img.SetPixelsToRange(DestRange)
    Call DestRange.Select
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
        Call RGBQuadToCell(objCell, CellToRGBQuad(objCell), True)
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
    If CheckSelection <> E_Range Then
        Call MsgBox("�Ώۂ͈̔͂�I�����Ă�����s���Ă�������")
        Exit Sub
    End If
    Set objSelection = Selection
    If objSelection.Rows.Count = Rows.Count Or objSelection.Columns.Count = Columns.Count Then
        Call MsgBox("���ׂĂ̍s�܂��͗�̑I�����͎��s�o���܂���")
        Exit Sub
    End If
    If objSelection.Count = 1 Then
        Call MsgBox("�Ώۂ͈̔͂�I�����Ă�����s���Ă�������")
        Exit Sub
    End If
    
    Dim objCell As Range
    Dim SrcColor As Long
    Set objCell = SelectCell("�u���O�̐F�̃Z����I�����Ă�������", ActiveCell)
    If objCell Is Nothing Then
        Exit Sub
    End If
    If Intersect(objSelection, objCell) Is Nothing Then
        Call MsgBox("�Z�����A�Ώ۔͈͂̓����ɂ���܂���")
        Exit Sub
    End If
    SrcColor = CellToRGBQuad(objCell(1))
    
    Dim DstColor As Long
    Set objCell = SelectCell("�u����̐F�̃Z����I�����Ă�������", ActiveCell)
    If objCell Is Nothing Then
        Exit Sub
    End If
    DstColor = CellToRGBQuad(objCell(1))
    
    '�I��͈͂̏d����r��
    Dim objCanvas As Range
    Set objCanvas = ReSelectRange(objSelection)
    
    '�u���ΏۃZ�����擾
    Dim objRange As Range
    For Each objCell In objCanvas
        If SrcColor = CellToRGBQuad(objCell) Then
            Set objRange = UnionRange(objRange, objCell)
        End If
    Next
    If objRange Is Nothing Then Exit Sub
    
    Application.ScreenUpdating = False
    Call SaveUndoInfo(Selection)
    Call RGBQuadToCell(objRange, DstColor, True)
    Call SetOnUndo("�F�̒u��")
    Call objSelection.Select
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
    If CheckSelection <> E_Range Then
        Call MsgBox("�Ώۂ͈̔͂�I�����Ă�����s���Ă�������")
        Exit Sub
    End If
    Set objSelection = Selection
    If objSelection.Rows.Count = Rows.Count Or objSelection.Columns.Count = Columns.Count Then
        Call MsgBox("���ׂĂ̍s�܂��͗�̑I�����͎��s�o���܂���")
        Exit Sub
    End If
    If objSelection.Count = 1 Then
        Call MsgBox("�Ώۂ͈̔͂�I�����Ă�����s���Ă�������")
        Exit Sub
    End If
    
    Dim SelectColor  As Long
    Dim strMsg As String
    If blnSameColor Then
        strMsg = "�I���������F�̃Z����I�����Ă�������"
    Else
        strMsg = "�I���������Ȃ��F�̃Z����I�����Ă�������"
    End If
    Dim objCell As Range
    Set objCell = SelectCell(strMsg, ActiveCell)
    If objCell Is Nothing Then
        Exit Sub
    End If
'    If Intersect(objSelection, objCell) Is Nothing Then
'        Call MsgBox("�Z�����A�Ώ۔͈͂̓����ɂ���܂���")
'        Exit Sub
'    End If
    SelectColor = CellToRGBQuad(objCell(1))
    
    '�I��͈͂̏d����r��
    Dim objCanvas As Range
    Set objCanvas = ReSelectRange(objSelection)
    
    Dim objNewSelection As Range
    For Each objCell In objCanvas
        If (SelectColor = CellToRGBQuad(objCell)) = blnSameColor Then
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
        Call ReSelectRange(objRange).Select
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
    Dim objCell As Range
    Dim Color   As Long
    If objCopyRange.Count > 1 Then
        If objSelection.Areas.Count > 1 Then
            Call MsgBox("���̃R�}���h�͕����̑I��͈͂ɑ΂��Ď��s�ł��܂���")
            Exit Sub
        End If
        If GetTmpControl("C1").State Then
            Set objCell = SelectCell("�ΏېF�̃Z����I�����Ă�������", ActiveCell)
            If objCell Is Nothing Then Exit Sub
            If Intersect(objCopyRange, objCell) Is Nothing Then
                Call MsgBox("�Z�����A�Ώ۔͈͂̓����ɂ���܂���")
                Exit Sub
            End If
            Color = CellToRGBQuad(objCell)
        End If
        If GetTmpControl("C2").State Then
            Set objCell = SelectCell("���O�Ώۂ̐F�̃Z����I�����Ă�������", ActiveCell)
            If objCell Is Nothing Then Exit Sub
            If Intersect(objCopyRange, objCell) Is Nothing Then
                Call MsgBox("�Z�����A�Ώ۔͈͂̓����ɂ���܂���")
                Exit Sub
            End If
            Color = CellToRGBQuad(objCell)
        End If
    End If
    
    '�\�t����̗̈�
    Dim objDestRange As Range
    If objCopyRange.Count = 1 Then
        Set objDestRange = objSelection
    Else
        Set objDestRange = objSelection.Resize(objCopyRange.Rows.Count, objCopyRange.Columns.Count)
    End If
    
    Application.ScreenUpdating = False
    Call SaveUndoInfo(objDestRange)
    If objCopyRange.Count = 1 Then
        Dim DstColor As Long
        DstColor = CellToRGBQuad(objCopyRange)
        Call RGBQuadToCell(objDestRange, DstColor, True)
    Else
        If GetTmpControl("C1").State Or GetTmpControl("C2").State Then
            Call PasteSub(Color, objCopyRange, objDestRange)
        Else
            Dim img As New CImage
            Call img.GetPixelsFromRange(objCopyRange)
            Call img.SetPixelsToRange(objDestRange)
        End If
    End If
    Call objDestRange.Select
    Call SetOnUndo("�\�t��")
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub


'*****************************************************************************
'[�T�v] ����̐F���� �܂��� ����̐F�������\�t��
'[����] �Ώۂ܂��͑ΏۊO�Ƃ���F,Copy����Range,�\�t�����Range
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub PasteSub(ByVal Color As Long, ByRef objCopyRange As Range, ByRef objDestRange As Range)
    Dim objSameRange  As Range  '�����F�̃Z��
    Dim objDiffRange  As Range  '�Ⴄ�F�̃Z��
    Dim objCell As Range
    
    '�I��F�Ɠ����F�̃Z���ƈႤ�F�̃Z�����擾
    Dim i As Long
    For Each objCell In objCopyRange
        i = i + 1
        If Color = CellToRGBQuad(objCell) Then
            Set objSameRange = UnionRange(objSameRange, objDestRange(i))
        Else
            Set objDiffRange = UnionRange(objDiffRange, objDestRange(i))
        End If
    Next
    
    '�X�V�ΏۃZ�����������̂��߂ɃN���A
    If GetTmpControl("C3").State Then
        '�\�t����̈�S�̂��N���A
        Call ClearRange(objDestRange)
    Else
        '�ΏۊO�̃Z���͍X�V���Ȃ���
        If GetTmpControl("C1").State Then
            '�����F�̃Z�����N���A
            Call ClearRange(objSameRange)
        Else
            '�Ⴄ�F�̃Z�����N���A
            Call ClearRange(objDiffRange)
        End If
    End If
    
    '�����Z���̐ݒ�
    If GetTmpControl("C3").State Then
        If GetTmpControl("C1").State Then
            '�Ⴄ�F�̃Z���𓧖���
            Call RGBQuadToCell(objDiffRange, CTRANSPARENT, False)
        Else
            '�����F�̃Z���𓧖���
            Call RGBQuadToCell(objSameRange, CTRANSPARENT, False)
        End If
    End If
    
    '�J���[�̐ݒ�
    If GetTmpControl("C1").State Then
        '�����F�̃Z����ݒ�
        Call RGBQuadToCell(objSameRange, Color, False)
    Else
        '�Ⴄ�F�̃Z����ݒ�
        i = 0
        For Each objCell In objCopyRange
            i = i + 1
            If Color <> CellToRGBQuad(objCell) Then
                Call RGBQuadToCell(objDestRange(i), CellToRGBQuad(objCell), False)
            End If
        Next
    End If
End Sub


'*****************************************************************************
'[�T�v] �h��Ԃ�
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub �h��Ԃ�()
On Error GoTo ErrHandle
    Dim objSelection As Range
    If CheckSelection <> E_Range Then
        Call MsgBox("�Ώۂ͈̔͂�I�����Ă�����s���Ă�������")
        Exit Sub
    End If
    Set objSelection = Selection
    If objSelection.Rows.Count = Rows.Count Or objSelection.Columns.Count = Columns.Count Then
        Call MsgBox("���ׂĂ̍s�܂��͗�̑I�����͎��s�o���܂���")
        Exit Sub
    End If
    If objSelection.Count = 1 Or objSelection.Areas.Count > 1 Then
        Call MsgBox("�Ώۂ͈̔͂�I�����Ă�����s���Ă�������")
        Exit Sub
    End If
    
    Dim objStartCell As Range
    Dim SrcColor As Long
    Set objStartCell = SelectCell("�h��Ԃ��J�n�Z����I�����Ă�������", ActiveCell)
    If objStartCell Is Nothing Then
        Exit Sub
    End If
    If Intersect(objSelection, objStartCell) Is Nothing Then
        Call MsgBox("�Z�����A�Ώ۔͈͂̓����ɂ���܂���")
        Exit Sub
    End If
    SrcColor = CellToRGBQuad(objStartCell)
    
    Dim objColorCell As Range
    Dim DstColor As Long
    Set objColorCell = SelectCell("�h��Ԃ��F�̃Z����I�����Ă�������", ActiveCell)
    If objColorCell Is Nothing Then
        Exit Sub
    End If
    DstColor = CellToRGBQuad(objColorCell)
    
    If SrcColor = DstColor Then Exit Sub

    Dim img As New CImage
    Call img.GetPixelsFromRange(objSelection)
    Call img.Fill(objStartCell.Column - objSelection.Column + 1, _
                  objStartCell.Row - objSelection.Row + 1, _
                  DstColor)
    Application.ScreenUpdating = False
    Call SaveUndoInfo(objSelection)
    Call img.SetPixelsToRange(objSelection)
    Call SetOnUndo("�h��Ԃ�")
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
Public Sub �F����(ByVal Up As Long)
On Error GoTo ErrHandle
    If CheckSelection <> E_Range Then Exit Sub
    If GetTmpControl("C4").State Or GetTmpControl("C5").State Or _
       GetTmpControl("C6").State Or GetTmpControl("C7").State Then
    Else
        Call MsgBox("RGB����уA���t�@�l�̂�������`�F�b�N����Ă��܂���")
        Exit Sub
    End If
    If Selection.Rows.Count = Rows.Count Or Selection.Columns.Count = Columns.Count Then
        Call MsgBox("���ׂĂ̍s�܂��͗�̑I�����͎��s�o���܂���")
        Exit Sub
    End If

    'Ctrl������
    If GetKeyState(vbKeyControl) < 0 Then
    Else
        Up = Up * 10
    End If
    
    '�Z���̑I����Ԃ�����(�}�`�I���̏�Ԃɂ���)
    Call ReleaseCellSelect

    '����R�}���h���A�ł���Ă��邩
    Dim IsBeat As Boolean
    If "�F����" = GetUndoStr() Then
        IsBeat = (RangeToAddress(Selection) = FSelection)
    End If
    
    'RGB���̑����l
    Static R As Long
    Static G As Long
    Static B As Long
    Static A As Long
    
    '����R�}���h���A�ł���Ă��邩
    If Not IsBeat Then
        R = 0
        G = 0
        B = 0
        A = 0
    End If
    If GetTmpControl("C4").State Then
        R = R + Up
    End If
    If GetTmpControl("C5").State Then
        G = G + Up
    End If
    If GetTmpControl("C6").State Then
        B = B + Up
    End If
    If GetTmpControl("C7").State Then
        A = A + Up
    End If
    
    Application.ScreenUpdating = False
    If Not IsBeat Then
        Call SaveUndoInfo(Selection, "�F����")
    End If
    
    Dim RGBQuad As Long
    Dim objCell As Range
    For Each objCell In ReSelectRange(Selection)
        RGBQuad = AdjustColor(CellToRGBQuad(ThisWorkbook.Worksheets(UndoSheetName).Range(objCell.Address)), R, G, B, A)
        Call RGBQuadToCell(objCell, RGBQuad, True)
    Next
    Call Selection.Select
    Call SetOnUndo("�F����")
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[�T�v] �Z���̑I����Ԃ�����(�}�`�I���̏�Ԃɂ���)
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub ReleaseCellSelect()
    On Error Resume Next
    If Not CommandBars.FindControl(, 182).State Then
        Call CommandBars.FindControl(, 182).Execute
    End If
End Sub

'*****************************************************************************
'[�T�v] �F��HSL�̒l�𑝌�������
'[����] 1:�����A-1:����
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub HSL����(ByVal Up As Long, ByVal lngType As Long)
On Error GoTo ErrHandle
    If CheckSelection <> E_Range Then Exit Sub
    If Selection.Rows.Count = Rows.Count Or Selection.Columns.Count = Columns.Count Then
        Call MsgBox("���ׂĂ̍s�܂��͗�̑I�����͎��s�o���܂���")
        Exit Sub
    End If

    'Ctrl������
    If GetKeyState(vbKeyControl) < 0 Then
    Else
        Up = Up * 5
    End If
    
    '�Z���̑I����Ԃ�����(�}�`�I���̏�Ԃɂ���)
    Call ReleaseCellSelect
    
    Dim strUndo As String
    Select Case lngType
    Case 1 '�F��
        strUndo = "�F��"
    Case 2 '�ʓx
        strUndo = "�ʓx"
    Case 3 '���x
        strUndo = "���x"
    End Select
    
    '����R�}���h���A�ł���Ă��邩
    Dim IsBeat As Boolean
    If strUndo = GetUndoStr() Then
        IsBeat = (RangeToAddress(Selection) = FSelection)
    End If
    
    Static H As Long
    Static S As Long
    Static L As Long
    
    '����R�}���h���A�ł���Ă��邩
    If Not IsBeat Then
        H = 0
        S = 0
        L = 0
    End If
    Select Case lngType
    Case 1 '�F��
        H = H + Up
    Case 2 '�ʓx
        S = S + Up
    Case 3 '���x
        L = L + Up
    End Select
    
    Application.ScreenUpdating = False
    If Not IsBeat Then
        Call SaveUndoInfo(Selection, strUndo)
    End If
    
    Dim RGBQuad As Long
    Dim objCell As Range
    For Each objCell In ReSelectRange(Selection)
        RGBQuad = UpDownHSL(CellToRGBQuad(ThisWorkbook.Worksheets(UndoSheetName).Range(objCell.Address)), H, S, L)
        Call RGBQuadToCell(objCell, RGBQuad, True)
    Next
    
    Call Selection.Select
    Call SetOnUndo(strUndo)
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
'[�T�v] �F�𐔒l��
'[����] True:RGBA(16�i8��),False:RGB(16�i6��)
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub �F�𐔒l��(ByVal blnAlpha As Boolean)
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
        If blnAlpha Then
            objCell.Value = Cell2RGBA(objCell)
        Else
            objCell.Value = Cell2RGB(objCell)
        End If
    Next
    
    '�t�H���g�̐F�ƖԊ|����W���ɖ߂�
    With objSelection
        .Font.ColorIndex = xlAutomatic
        .Interior.Pattern = xlSolid
    End With
    Call Selection.Select
    Call SetOnUndo("�F�𐔒l��")
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub


'*****************************************************************************
'[�T�v] ���l����F��ݒ�
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub ���l����F��ݒ�()
On Error GoTo ErrHandle
    If CheckSelection <> E_Range Then Exit Sub
    If Selection.Rows.Count = Rows.Count Or Selection.Columns.Count = Columns.Count Then
        Call MsgBox("���ׂĂ̍s�܂��͗�̑I�����͎��s�o���܂���")
        Exit Sub
    End If
        
    '�l�̓��͂��ꂽ�Z���̂ݑΏ�
    Dim objSelection As Range
    If Selection.Count <> 1 Then
        Dim objCells(1 To 3) As Range
        With Selection
            On Error Resume Next
            Set objCells(1) = .SpecialCells(xlCellTypeConstants)
            Set objCells(2) = .SpecialCells(xlCellTypeFormulas)
            On Error GoTo 0
        End With
        Set objSelection = UnionRange(objCells(1), objCells(2))
        If objSelection Is Nothing Then Exit Sub
    Else
        Set objSelection = Selection
    End If
    
    '�ΏۃZ�����擾
    Dim objRange As Range
    Dim objZero As Range
    Dim objCell As Range
    Dim vValue  As Variant
    For Each objCell In objSelection
        vValue = objCell.Value
        If IsNumeric(vValue) Then
            If vValue = 0 Then
                Set objZero = UnionRange(objZero, objCell)
            End If
        ElseIf Left(vValue, 1) = "#" Then
            If Len(vValue) = 7 Or Len(vValue) = 9 Then
                If IsNumeric("&H" & Mid(vValue, 2)) Then
                   Set objRange = UnionRange(objRange, objCell)
                End If
            End If
        End If
    Next
    If (objRange Is Nothing) And (objZero Is Nothing) Then Exit Sub
        
    Application.ScreenUpdating = False
    Call SaveUndoInfo(Selection)
    
    '�����F�ȊO
    If Not (objRange Is Nothing) Then
        '�������̂��ߏ������N���A
        With objRange
            .Interior.Pattern = xlNone
            .Font.Color = xlAutomatic
        End With
        
        Dim RGBQuad As TRGBQuad
        For Each objCell In objRange
            vValue = objCell.Value
            With RGBQuad
                .Red = "&H" & Mid(vValue, 2, 2)
                .Green = "&H" & Mid(vValue, 4, 2)
                .Blue = "&H" & Mid(vValue, 6, 2)
                If Len(vValue) = 9 Then
                    '8���̎�
                    .Alpha = "&H" & Mid(vValue, 8, 2)
                Else
                    '6���̎��A�s����
                    .Alpha = 255
                End If
            End With
            Dim Color As TLong
            LSet Color = RGBQuad
            Call RGBQuadToCell(objCell, Color.Long, True)
        Next
    End If
    
    '�����F
    If Not (objZero Is Nothing) Then
        Call RGBQuadToCell(objZero, CTRANSPARENT, True)
    End If
    Call Selection.Select
    Call SetOnUndo("���l����F��ݒ�")
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[�T�v] �A���t�@�l��\��
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub �A���t�@�l��\��()
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
    
On Error Resume Next
    Dim Alpha As Byte
    Dim vValue As String
    Dim objCell As Range
    For Each objCell In objSelection
        With objCell.Interior
            Select Case .ColorIndex
            Case xlNone, xlAutomatic
                '����
                Alpha = 0
            Case Else
                '�s����
                Alpha = 255
                '���������ǂ���
                If .Pattern = xlGray8 Then
                    vValue = objCell.Value
                    If IsNumeric(vValue) Then
                        If 0 <= CLng(vValue) And CLng(vValue) <= 255 Then
                            '�Z���ɓ��͂��ꂽ���l���A���t�@�l
                            Alpha = vValue
                        End If
                    End If
                End If
            End Select
        End With
        objCell.Value = Alpha
    Next
On Error GoTo ErrHandle
    
    '�t�H���g�̐F��W���ɖ߂�
    With objSelection.Font
        .ColorIndex = xlAutomatic
    End With
    Call SetOnUndo("���l�\��")
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[�T�v] ���l����A���t�@�l��ݒ�
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub ���l����A���t�@�l��ݒ�()
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
    
On Error Resume Next
    '�ΏۃZ�����擾
    Dim objRange As Range
    Dim objZero As Range
    Dim obj255 As Range
    Dim objCell As Range
    Dim vValue  As Variant
    For Each objCell In objSelection
        vValue = objCell.Value
        If IsNumeric(vValue) And vValue <> "" Then
            Select Case vValue
            Case 0 To 255
                Select Case objCell.Interior.ColorIndex
                Case xlNone, xlAutomatic
                    Set objZero = UnionRange(objZero, objCell)
                Case Else
                    Select Case vValue
                    Case 0
                        Set objZero = UnionRange(objZero, objCell)
                    Case 255
                        Set obj255 = UnionRange(obj255, objCell)
                    Case 1 To 254
                        Set objRange = UnionRange(objRange, objCell)
                    End Select
                End Select
            End Select
        End If
    Next
    If (objRange Is Nothing) And (objZero Is Nothing) And (obj255 Is Nothing) Then Exit Sub
        
    Application.ScreenUpdating = False
    Call SaveUndoInfo(Selection)
    
    '�����F
    If Not (objZero Is Nothing) Then
        Call RGBQuadToCell(objZero, CTRANSPARENT, True)
    End If
    
    '�s�����F
    If Not (obj255 Is Nothing) Then
        With obj255
            .Interior.Pattern = xlAutomatic
            .Font.Color = xlAutomatic
            .ClearContents
        End With
    End If

    '������
    If Not (objRange Is Nothing) Then
        With objRange.Interior
            .Pattern = xlGray8
            .PatternColor = &HFFFFFF '��
        End With
        
        For Each objCell In objRange
            With objCell.Interior
                objCell.Font.Color = .Color '������w�i�F�Ɠ����ɂ���
            End With
        Next
    End If
    
    Call Selection.Select
    Call SetOnUndo("���l�ݒ�")
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

