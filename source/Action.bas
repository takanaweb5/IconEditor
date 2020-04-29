Attribute VB_Name = "Action"
Option Explicit

Private Declare PtrSafe Function GetKeyState Lib "user32" (ByVal lngVirtKey As Long) As Integer
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hDC As LongPtr) As Long

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As LongPtr) As LongPtr
Private Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal hDC As LongPtr) As Long
Private Declare PtrSafe Function GetDIBits Lib "gdi32" (ByVal aHDC As LongPtr, ByVal hBitmapptr As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, ByRef lpBits As Any, ByRef lpBI As Any, ByVal wUsage As Long) As Long
Private Const DIB_RGB_COLORS = 0&

Private Declare PtrSafe Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As LongPtr, ByVal nCount As Long, lpObject As Any) As Long
Private Type BITMAPINFOHEADER
    Size          As Long
    Width         As Long
    Height        As Long
    Planes        As Integer
    BitCount      As Integer
    Compression   As Long
    SizeImg     As Long
    XPelsPerMeter As Long
    YPelsPerMeter As Long
    ClrUsed       As Long
    ClrImportant  As Long
End Type

Private Type BITMAPINFO
    Header As BITMAPINFOHEADER
    Colors(1 To 256) As Long
End Type

Private Type TBITMAP
    Type As Long
    Width As Long
    Height As Long
    WidthBytes As Long
    Planes As Integer
    BitsPixel As Integer
#If Win64 Then
    Bits As LongPtr
    Reserve As Long '�\���̂�64Bit�̔{���ɂ��邽��
#Else
    Bits As Long
#End If
End Type

Private Const LOGPIXELSX = 88
Private Const LOGPIXELSY = 90

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
        strInput = InputBox("��,��������͂��Ă�������" & vbCrLf & vbCrLf & "�� 32,32")
        If strInput = "" Then
            lngWidth = 32
            lngHeight = 32
            Exit Do
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
    Call img.SetPixelsToRange(ActiveCell)
    Call ActiveCell.Resize(img.Height, img.Width).Select
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
    Call img.LoadFromFile(vDBName)
    If img.Width > 256 Or img.Height > 256 Then
        Call MsgBox("���܂��͍�����256Pixel�𒴂���t�@�C���͓ǂݍ��߂܂���")
        Exit Sub
    End If
    Application.ScreenUpdating = False
    Call img.SetPixelsToRange(ActiveCell)
    Call ActiveCell.Resize(img.Height, img.Width).Select
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
    If Selection.Rows.Count > 256 Or Selection.Columns.Count > 256 Then
        Call MsgBox("���܂��͍�����256Pixel�𒴂���摜�͕ۑ��ł��܂���")
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
    Dim img As New CImage
    Call img.GetPixelsFromRange(Selection)
    Call img.FlipHorizontal
    Application.ScreenUpdating = False
    Call img.SetPixelsToRange(Selection)
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
    Dim img As New CImage
    Call img.GetPixelsFromRange(Selection)
    Call img.FlipVertical
    Application.ScreenUpdating = False
    Call img.SetPixelsToRange(Selection)
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
    Call img.SetPixelsToRange(Selection)
    Call ActiveCell.Resize(img.Height, img.Width).Select
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
            lngWidth = img.Width
            lngHeight = img.Height
            Exit Do
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
    Call img.SetPixelsToRange(Selection)
    Call ActiveCell.Resize(img.Height, img.Width).Select
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
    Dim img As New CImage
    If CheckSelection <> E_Range Then Exit Sub
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
        strInput = InputBox("��,������64Pixel�����œ��͂��Ă�������")
        If strInput = "" Then
            Exit Do
        End If
        WidthAndHeight = Split(strInput, ",")
        If UBound(WidthAndHeight) = 1 Then
            If IsNumeric(WidthAndHeight(0)) And IsNumeric(WidthAndHeight(1)) Then
                lngWidth = WidthAndHeight(0)
                lngHeight = WidthAndHeight(1)
                If lngWidth <= 64 And lngHeight <= 64 Then
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
    Call img.SetPixelsToRange(objCell)
    Call objCell.Resize(img.Height, img.Width).Select
    If Not (objWkShape Is Nothing) Then
        Call objWkShape.Delete
    End If
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
    
    Dim objSelection As Range
    Dim objCell As Range
    Set objSelection = Selection
    If objSelection.Rows.Count > 256 Or objSelection.Columns.Count > 256 Then
        Call MsgBox("���܂��͍�����256�}�X�𒴂��鎞�͎��s�o���܂���")
        Exit Sub
    End If
    
    '�I��͈͂̏d����r��
    Set objSelection = ReSelectRange(objSelection)
    
    Dim img As New CImage
    Dim objArea As Range
    For Each objArea In objSelection.Areas
        Call img.GetPixelsFromRange(objArea)
        Call img.SetPixelsToRange(objArea)
    Next
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
        If objCanvas.Rows.Count > 256 Or objCanvas.Columns.Count > 256 Then
            Call MsgBox("���܂��͍�����256�}�X�𒴂��鎞�͎��s�o���܂���")
            Exit Sub
        End If
    End If
    
    Dim objSrcCell As Range
    Set objSrcCell = SelectCell("�u���O�̐F�̃Z����I�����Ă�������", ActiveCell)
    If objSrcCell Is Nothing Then
        Exit Sub
    End If
    Dim objDstCell As Range
    Set objDstCell = SelectCell("�u����̐F�̃Z����I�����Ă�������", ActiveCell)
    If objDstCell Is Nothing Then
        Exit Sub
    End If
    
    '�I��͈͂̏d����r��
    Set objCanvas = ReSelectRange(objCanvas)
    
    Dim img As New CImage
    Dim lngColor As Long
    lngColor = img
    Dim objArea As Range
    For Each objArea In objCanvas
        If img.SameCellColor(objSrcCell(1), objArea) Then
            Debug.Print objArea.Address(0, 0)
            Call img.ChangeCellColor(objDstCell(1), objArea)
        End If
    Next
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
        If objCanvas.Rows.Count > 256 Or objCanvas.Columns.Count > 256 Then
            Call MsgBox("���܂��͍�����256�}�X�𒴂��鎞�͎��s�o���܂���")
            Exit Sub
        End If
    End If
    
    Dim objColorCell As Range
    Dim strMsg As String
    If blnSameColor Then
        strMsg = "�I���������F�̃Z����I�����Ă�������"
    Else
        strMsg = "�I���������Ȃ��F�̃Z����I�����Ă�������"
    End If
    Set objColorCell = SelectCell(strMsg, ActiveCell)
    If objColorCell Is Nothing Then
        Exit Sub
    End If
    
    Dim img As New CImage
    Dim objNewSelection As Range
    Dim objCell As Range
    For Each objCell In objCanvas
        If img.SameCellColor(objColorCell, objCell) = blnSameColor Then
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
    Dim objCopyRange  As Range
    Set objCopyRange = GetCopyRange()
    If objCopyRange Is Nothing Then Exit Sub
    If objCopyRange.Rows.Count > 256 Or objCopyRange.Columns.Count > 256 Then
        Call MsgBox("���܂��͍�����256�}�X�𒴂��鎞�͎��s�o���܂���")
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
    
    Dim img As New CImage
    Application.ScreenUpdating = False
    Call img.GetPixelsFromRange(objCopyRange)
    If objCopyRange.Count > 1 Then
        Call img.SetPixelsToRange(objSelection, lngMode, objColorCell, FChecked(3))
    Else
        Dim objCell As Range
        For Each objCell In objSelection
            Call img.SetPixelsToRange(objCell)
        Next
    End If
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
    If Selection.Rows.Count > 256 Or Selection.Columns.Count > 256 Then
        Call MsgBox("���܂��͍�����256�}�X�𒴂��鎞�͎��s�o���܂���")
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
    
    Dim img As New CImage
    Dim objArea As Range
    For Each objArea In objSelection.Areas
        Call img.GetPixelsFromRange(objArea)
        Call img.AdjustColor(lngUp, FChecked(4), FChecked(5), FChecked(6), FChecked(7))
        Call img.SetPixelsToRange(objArea)
    Next
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
        If objCanvas.Rows.Count > 256 Or objCanvas.Columns.Count > 256 Then
            Call MsgBox("���܂��͍�����256�}�X�𒴂��鎞�͎��s�o���܂���")
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
    Call img.SetPixelsToRange(objCanvas)
    
Exit Sub
ErrHandle:
    Call MsgBox(Err.Description, vbExclamation)
End Sub

'*****************************************************************************
'[�T�v] �̈�Ɨ̈�̏d�Ȃ�̈���擾����
'[����] �Ώۗ̈�(Nothing����)
'[�ߒl] objRange1 �� objRange2
'*****************************************************************************
Private Function IntersectRange(ByRef objRange1 As Range, ByRef objRange2 As Range) As Range
    Select Case True
    Case (objRange1 Is Nothing) Or (objRange2 Is Nothing)
        Set IntersectRange = Nothing
    Case Else
        Set IntersectRange = Intersect(objRange1, objRange2)
    End Select
End Function

'*****************************************************************************
'[�T�v] �̈�ɗ̈��������
'[����] �Ώۗ̈�(Nothing����)
'[�ߒl] objRange1 �� objRange2
'*****************************************************************************
Private Function UnionRange(ByRef objRange1 As Range, ByRef objRange2 As Range) As Range
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
'[�T�v] �̈悩��̈���A���O����
'       �`�|�a = �`��!�a
'       !�a = !(B1��B2��B3...��Bn) = !B1��!B2��!B3...��!Bn
'[����] �Ώۗ̈�
'[�ߒl] objRange1 �| objRange2
'*****************************************************************************
Private Function MinusRange(ByRef objRange1 As Range, ByRef objRange2 As Range) As Range
    Dim objRounds As Range
    Dim i As Long
    
    If objRange2 Is Nothing Then
        Set MinusRange = objRange1
        Exit Function
    End If
    
    '���O����̈�̐��������[�v
    '!�a = !B1��!B2��!B3.....��!Bn
    Set objRounds = ReverseRange(objRange2.Areas(1))
    For i = 2 To objRange2.Areas.Count
        Set objRounds = IntersectRange(objRounds, ReverseRange(objRange2.Areas(i)))
    Next
    
    '�`��!�a
    Set MinusRange = IntersectRange(objRange1, objRounds)
End Function

'*****************************************************************************
'[�T�v] �̈�𔽓]����
'[����] �Ώۗ̈�
'[�ߒl] !objRange
'*****************************************************************************
Private Function ReverseRange(ByRef objRange As Range) As Range
    Dim i As Long
    Dim objRound(1 To 4) As Range
    
    With objRange.Parent
        On Error Resume Next
        '�I��̈����̗̈悷�ׂ�
        Set objRound(1) = .Range(.Rows(1), _
                                 .Rows(objRange.Row - 1))
        '�I��̈��艺�̗̈悷�ׂ�
        Set objRound(2) = .Range(.Rows(objRange.Row + objRange.Rows.Count), _
                                 .Rows(Rows.Count))
        '�I��̈��荶�̗̈悷�ׂ�
        Set objRound(3) = .Range(.Columns(1), _
                                 .Columns(objRange.Column - 1))
        '�I��̈���E�̗̈悷�ׂ�
        Set objRound(4) = .Range(.Columns(objRange.Column + objRange.Columns.Count), _
                                 .Columns(Columns.Count))
        On Error GoTo 0
    End With
    
    '�I��̈�ȊO�̗̈��ݒ�
    For i = 1 To 4
        Set ReverseRange = UnionRange(ReverseRange, objRound(i))
    Next
End Function

'*****************************************************************************
'[�T�v] �̈�̏d�����Ȃ����̈���擾
'[����] �Ώۗ̈�
'[�ߒl] �̈�̏d�����Ȃ����̈�
'*****************************************************************************
Private Function ReSelectRange(ByRef objRange As Range) As Range
    Dim objArrange(1 To 3) As Range
    With objRange
        On Error Resume Next
        Set objArrange(1) = .SpecialCells(xlCellTypeConstants)
        Set objArrange(2) = .SpecialCells(xlCellTypeFormulas)
        Set objArrange(3) = .SpecialCells(xlCellTypeBlanks)
        On Error GoTo 0
    End With

    Dim i As Long
    For i = 1 To 3
        Set ReSelectRange = UnionRange(ReSelectRange, objArrange(i))
    Next
End Function

