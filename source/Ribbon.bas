Attribute VB_Name = "Ribbon"
Option Explicit
Option Private Module

Private Const PAGE_READONLY = 2&
Private Const PAGE_READWRITE = 4&
Private Const FILE_MAP_WRITE = 2&
Private Const FILE_MAP_READ = 4&

Private Declare PtrSafe Function CreateFileMapping Lib "kernel32" Alias "CreateFileMappingW" (ByVal hFile As LongPtr, lpFileMappingAttributes As Any, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As String) As LongPtr
Private Declare PtrSafe Function OpenFileMapping Lib "kernel32" Alias "OpenFileMappingW" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal ptrToNameString As String) As LongPtr
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function MapViewOfFile Lib "kernel32" (ByVal hFileMappingObject As LongPtr, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As LongPtr
Private Declare PtrSafe Function UnmapViewOfFile Lib "kernel32" (ByVal lpBaseAddress As LongPtr) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)

'Private FRibbon As IRibbonUI '��O�����N���Ă��l�����Ȃ��Ȃ��悤�ɋ��L�������ɕύX
'Public FChecked(1 To 7) As Boolean
Private FSampleClick As Boolean

'*****************************************************************************
'[�T�v] IRibbonUI��ۑ�����CommandBar���쐬����
'       ���킹�āA���{���R���g���[���̏�Ԃ�ۑ�����CommandBarControl���쐬����
'       ���W���[���ϐ��ɕۑ������ꍇ�́A���R���p�C����R�[�h�̋�����~�Œl�����Ȃ��邽��
'[����] IRibbonUI
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub CreateTmpCommandBar(ByRef Ribbon As IRibbonUI)
    On Error Resume Next
    Call Application.CommandBars(ThisWorkbook.Name).Delete
    On Error GoTo 0
    
    Dim i As Long
    Dim objCmdBar As CommandBar
    Set objCmdBar = CommandBars.Add(ThisWorkbook.Name, Position:=msoBarPopup, Temporary:=True)
    With objCmdBar.Controls.Add(msoControlButton)
        .Tag = "RibbonUI" & ThisWorkbook.Name
        .Parameter = ObjPtr(Ribbon)
    End With
    
    '�`�F�b�N�{�b�N�X�̃N���[�����e���|�����ɍ쐬
    For i = 1 To 7
        With objCmdBar.Controls.Add(msoControlButton)
            .Tag = "C" & i & ThisWorkbook.Name
            .State = False '�����ݒ�̓`�F�b�N�Ȃ�
        End With
    Next
    
    'RGB�{�^���̏����l���`�F�b�N����ɐݒ�
    GetTmpControl("C4").State = True 'Red�`�F�b�N�{�b�N�X
    GetTmpControl("C5").State = True 'Gereen�`�F�b�N�{�b�N�X
    GetTmpControl("C6").State = True 'Blue�`�F�b�N�{�b�N�X
End Sub

'*****************************************************************************
'[�T�v] CommandBar����IRibbonUI���擾����
'[����] �Ȃ�
'[�ߒl] IRibbonUI
'*****************************************************************************
Private Function GetRibbonUI() As IRibbonUI
    Dim Pointer  As LongPtr
    With CommandBars.FindControl(, , "RibbonUI" & ThisWorkbook.Name)
        Pointer = .Parameter
    End With
    Dim obj As Object
    Call CopyMemory(obj, Pointer, Len(Pointer))
    Set GetRibbonUI = obj
End Function

'*****************************************************************************
'[�C�x���g] onLoad
'*****************************************************************************
Sub onLoad(Ribbon As IRibbonUI)
    '���{��UI���e���|�����̃R�}���h�o�[�ɕۑ�����
    '(���W���[���ϐ��ɕۑ������ꍇ�́A��O��R�[�h�̋�����~�Œl�����Ȃ��邽��)
    Call CreateTmpCommandBar(Ribbon)
'    Set FRibbon = Ribbon
End Sub

'*****************************************************************************
'[�C�x���g] loadImage
'*****************************************************************************
Sub loadImage(imageID As String, ByRef returnedVal)
  returnedVal = imageID
End Sub

'*****************************************************************************
'[�C�x���g] getVisible
'*****************************************************************************
Sub getVisible(control As IRibbonControl, ByRef returnedVal)
    returnedVal = True
End Sub

'*****************************************************************************
'[�C�x���g] getEnabled
'*****************************************************************************
Sub getEnabled(control As IRibbonControl, ByRef returnedVal)
    Select Case control.ID
    Case "C3"
        returnedVal = (GetTmpControl("C1").State Or GetTmpControl("C2").State)
    Case Else
        returnedVal = True
    End Select
End Sub

'*****************************************************************************
'[�C�x���g] getShowLabel
'*****************************************************************************
Sub getShowLabel(control As IRibbonControl, ByRef returnedVal)
    returnedVal = (GetTips(control, 1) <> "")
End Sub

'*****************************************************************************
'[�C�x���g] getLabel
'*****************************************************************************
Sub getLabel(control As IRibbonControl, ByRef returnedVal)
    returnedVal = GetTips(control, 1)
End Sub

'*****************************************************************************
'[�C�x���g] getScreentip
'*****************************************************************************
Sub getScreentip(control As IRibbonControl, ByRef returnedVal)
    returnedVal = GetTips(control, 2)
End Sub

'*****************************************************************************
'[�C�x���g] getSupertip
'*****************************************************************************
Sub getSupertip(control As IRibbonControl, ByRef returnedVal)
    returnedVal = GetTips(control, 3)
End Sub

'*****************************************************************************
'[�C�x���g] getShowImage
'*****************************************************************************
Sub getShowImage(control As IRibbonControl, ByRef returnedVal)
    Select Case Mid(control.ID, 2)
    Case Else
        returnedVal = True
    End Select
End Sub

'*****************************************************************************
'[�C�x���g] getImage
'*****************************************************************************
Sub getImage(control As IRibbonControl, ByRef returnedVal)
    Call GetImages(control, returnedVal)
End Sub

'*****************************************************************************
'[�C�x���g] getSize
'*****************************************************************************
Sub getSize(control As IRibbonControl, ByRef returnedVal)
    Select Case Mid(control.ID, 2)
    Case 11, 21, 32, 61, 62, 63
        returnedVal = 1
    Case Else
        returnedVal = 0
    End Select
End Sub

'*****************************************************************************
'[�C�x���g] getPressed
'*****************************************************************************
Sub getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = GetTmpControl(control.ID).State
End Sub

'*****************************************************************************
'[�C�x���g] onCheckAction
'*****************************************************************************
Sub onCheckAction(control As IRibbonControl, pressed As Boolean)
    '�`�F�b�N��Ԃ�ۑ�
    GetTmpControl(control.ID).State = pressed
    
    Select Case control.ID
    Case "C1"
'        Application.EnableEvents = False
        '����F�E����F�ȊO�̃g�O��
        GetTmpControl("C2").State = False
        GetTmpControl("C3").State = False
        Call GetRibbonUI.InvalidateControl("C2")
        
        '�L��������؂�ւ�
        Call GetRibbonUI.InvalidateControl("C3")
'        Application.EnableEvents = True
    Case "C2"
'        Application.EnableEvents = False
        '����F�E����F�ȊO�̃g�O��
        GetTmpControl("C1").State = False
        GetTmpControl("C3").State = False
        Call GetRibbonUI.InvalidateControl("C1")
        
        '�L��������؂�ւ�
        Call GetRibbonUI.InvalidateControl("C3")
'        Application.EnableEvents = True
    End Select
End Sub

'*****************************************************************************
'[�T�v] Label�����ScreenTip��ݒ肵�܂�
'[����] lngType�u1:getLabel, 2:getScreentip, 3:getSupertip�v
'[�ߒl] �ݒ�l
'*****************************************************************************
Private Function GetTips(control As IRibbonControl, ByVal lngType As Long) As String
    ReDim Result(1 To 3) '1:getLabel, 2:getScreentip, 3:getSupertip
    Select Case Mid(control.ID, 2)
    Case 11
        Result(1) = "ImageMso"
        Result(2) = "ImageMso����摜���擾"
        Result(3) = "ImageMso���w�肵�đI�����ꂽ�Z���͈͂։摜��ǂݍ��݂܂�" & vbCrLf & "�P��Z���̑I�����̓T�C�Y���w�肷��_�C�A���O���\������܂�"
    Case 12
        Result(1) = "�Ǎ�"
        Result(2) = "�摜�̓Ǎ�"
        Result(3) = "�I�����ꂽ�Z���͈͂։摜��ǂݍ��݂܂�" & vbCrLf & "�P��Z���̑I�����͌��̉摜�̃T�C�Y�œǂݍ��݂܂�"
    Case 13
        Result(1) = "�ꊇ�Ǎ�"
        Result(2) = "�摜�̈ꊇ�Ǎ���"
        Result(3) = "�ꊇ�Ǎ��ݗp�̃V�[�g���J���܂�"
    Case 14
        Result(1) = "�ۑ�"
        Result(2) = "�摜�̕ۑ�"
        Result(3) = "�I�����ꂽ�Z���͈͂̉摜��ۑ����܂�"
    Case 15
        Result(1) = "�ꊇ�ۑ�"
        Result(2) = "�摜�̈ꊇ�ۑ�"
        Result(3) = "�ꊇ�ۑ��p�̃V�[�g���J���܂�"
    Case 16
        Result(1) = "Clipbord�摜�ۑ�"
        Result(2) = "�N���b�v�{�[�h�̉摜��ۑ�"
        Result(3) = "�N���b�v�{�[�h�̉摜���t�@�C���ɕۑ����܂�"
    
    Case 21
        Result(1) = "�F��" & vbCrLf & "�\�t"
        Result(2) = "�\�t��"
        Result(3) = "�Z���̐F���̂ݓ\�t���܂�"
    Case 22
        Result(1) = "�摜��\�t"
        Result(2) = "�摜��\�t��"
        Result(3) = "�N���b�v�{�[�h�̉摜��I�����ꂽ�Z���͈͂֓\�t���܂�" & vbCrLf & "�P��Z���̑I�����͌��̉摜�̃T�C�Y�œ\�t���܂�"
    Case 23
        Result(1) = "Shape��\�t"
        Result(2) = "Shape��\�t��"
        Result(3) = "�I�[�g�V�F�C�v�̉摜��\�t���܂�"
    Case 24
        Result(1) = "90��"
        Result(2) = "�E��90����]���ē\�t��"
        Result(3) = "�̈���E��90����]���ē\�t���܂�"
    Case 25
        Result(1) = "-90��"
        Result(2) = "����90����]���ē\�t��"
        Result(3) = "�̈������90����]���ē\�t���܂�"
    
    Case 31
        Result(1) = "�N���A"
        Result(2) = "�N���A"
        Result(3) = "�I�����ꂽ�Z���͈͂̐F���N���A���܂�"
    Case 32
        Result(1) = "�����F�̋���"
        Result(2) = "�����F�̋���"
        Result(3) = "�I�����ꂽ�Z���͈͂�" & vbCrLf & "�����F�̃Z����Ԋ|���ŋ������A" & vbCrLf & "�ȊO�̂̃Z���͖Ԋ|���������܂�"
    Case 33
        Result(1) = "�h�ׂ�"
        Result(2) = "�h�ׂ�"
        Result(3) = "�I�����ꂽ�Z���͈͂�Ώۂɓh�ׂ��܂�"
    Case 34
        Result(1) = "�F�̒u��"
        Result(2) = "�F�̒u��"
        Result(3) = "�I�����ꂽ�Z���͈͂̐F��u�����܂�"
    Case 35
        Result(1) = ""
        Result(2) = "���E���]"
        Result(3) = "�I�����ꂽ�Z���͈͂����E���]���܂�"
    Case 36
        Result(1) = ""
        Result(2) = "�㉺���]"
        Result(3) = "�I�����ꂽ�Z���͈͂��㉺���]���܂�"
    Case 37
        Result(1) = ""
        Result(2) = "�E��90����]"
        Result(3) = "�I�����ꂽ�Z���͈͂��E��90����]���܂�"
    Case 38
        Result(1) = ""
        Result(2) = "����90����]"
        Result(3) = "�I�����ꂽ�Z���͈͂�����90����]���܂�"
    
    Case 41
        Result(1) = "�����F"
        Result(2) = "�����F�̃Z���̑I��"
        Result(3) = "�����F�̃Z����I�����܂�"
    Case 42
        Result(1) = "�Ⴄ�F"
        Result(2) = "�Ⴄ�F�̃Z���̑I��"
        Result(3) = "�Ⴄ�F�̃Z����I�����܂�"
    Case 43
        Result(1) = "���]��"
        Result(2) = "�I��̈�̔��]�Ȃ�"
        Result(3) = "�I��̈�̔��]��ꕔ���O�Ȃǂ��s���܂�"
    
    Case 51, 52
        Result(1) = ""
        Result(2) = "���邳"
        Result(3) = "�I�������͈͂̐F�̖��邳(0�`255)��1�P��(Ctrl��10�P��)�ŕύX���܂�" & vbCrLf & "����:�������F����" & vbCrLf & "����:�������F����"
    Case 53, 54
        Result(1) = ""
        Result(2) = "�ʂ₩��"
        Result(3) = "�I�������͈͂̐F�̍ʂ₩��(0�`255)��1�P��(Ctrl��10�P��)�ŕύX���܂�" & vbCrLf & "����:�D�F�����F" & vbCrLf & "����:�D�F�����F"
    Case 55, 56
        Result(1) = ""
        Result(2) = "�F��"
        Result(3) = "�I�������͈͂̐F��(0�`360��)��1���P��(Ctrl��5���P��)�ŕω������܂�" & vbCrLf & "����:�ԁ������΁���������" & vbCrLf & "����:�ԁ������΁���������"
    Case 57
        Result(1) = ""
        Result(2) = "RGB�e�F�̐��l(0�`255)����уA���t�@�l�����������܂�"
        Result(3) = "�I�������͈͂̃`�F�b�N����RGB�e�F�̐��l(0�`255)��1�P��(Ctrl��10�P��)�Ō��������܂�"
    Case 58
        Result(1) = ""
        Result(2) = "RGB�e�F�̐��l����уA���t�@�l�𑝉������܂�"
        Result(3) = "�I�������͈͂̃`�F�b�N����RGB�e�F�̐��l(0�`255)��1�P��(Ctrl��10�P��)�ő��������܂�"
    
    Case 61 To 66
        Select Case Mid(control.ID, 2)
        Case 61, 64
            Result(1) = "��1"
        Case 62, 65
            Result(1) = "��2"
        Case 63, 66
            Result(1) = "��3"
        End Select
        Select Case Mid(control.ID, 2)
        Case 61 To 63
            Result(2) = "�T���v���\��(��)"
        Case Else
            Result(2) = "�T���v���\��(��)"
        End Select
        Result(3) = "�I�����ꂽ�Z���͈͂̉摜��\�����܂�" & vbLf & _
                    "���킹�ăN���b�v�{�[�h�ɂ��摜���R�s�[���܂�"
    Case 71
        Result(1) = "�F�����l(RGB)"
        Result(2) = "�F�𐔒l��"
        Result(3) = "�F��6��(RGB)��16�i���ŃZ���ɐݒ肵�܂�" & vbLf & "�������������F��0�Ƃ��܂�"
    Case 72
        Result(1) = "�F�����l(RGBA)"
        Result(2) = "�F�𐔒l��"
        Result(3) = "�F��8��(RGBA)��16�i���ŃZ���ɐݒ肵�܂�" & vbLf & "�������������F��0�Ƃ��܂�"
    Case 73
        Result(1) = "�������l"
        Result(2) = "�Z���̏�Ԃ���A���t�@�l��\��"
        Result(3) = "���F�̃Z���ɂ́A0��\�����܂�" & vbLf & _
                    "�Z�����Ԋ|���Ń��l�����͂���Ă���Z���́A���l��\�����܂�" & vbLf & _
                    "�s�����̃Z����255��\�����܂�"
    Case 74
        Result(1) = "�Z���֐�"
        Result(2) = "�Z���֐�"
        Result(3) = "�Z���֐��̏Љ�ł�"
    Case 75
        Result(1) = "���l���F"
        Result(2) = "�Z���̒l����F��ݒ�"
        Result(3) = "6���܂���8����16�i���̃Z���̒l��F�ɕϊ����܂�" & vbLf & _
                    "��L�ɊY�����Ȃ��Z���́A�������܂���" & vbLf & _
                    "8���̎��͉E2�����A���t�@�l�ɐݒ肵�܂�" & vbLf & _
                    "6���̎��̓A���t�@�l(�����������)��ݒ肵�܂���" & vbLf & _
                    "���s��̓Z���̒l���N���A���܂�"
    Case 76
        Result(1) = "���l����"
        Result(2) = "�Z���̐��l����A���t�@�l��ݒ�"
        Result(3) = "0�̃Z���͓����ɂ��܂�" & vbLf & _
                    "1�`254�̃Z���̓A���t�@�l��ݒ肵�܂�" & vbLf & _
                    "�ȊO(255���)�̃Z���́A�s�����ɂ��܂�" & vbLf & _
                    "���s��̓Z���̒l���N���A���܂�"
    Case Else
        Result(1) = ""
    End Select
    
    GetTips = Result(lngType)
End Function

'*****************************************************************************
'[�T�v] Image��ݒ肵�܂�
'[����] control
'[�ߒl] Result
'*****************************************************************************
Private Sub GetImages(control As IRibbonControl, ByRef Result)
    Dim strImage As String
    Dim objImage As IPictureDisp

    Select Case Mid(control.ID, 2)
    Case 11
        strImage = "PictureReset"
    Case 12
        strImage = "FileOpen"
    Case 13
        strImage = "NewO12FilesTool"
    Case 14
        strImage = "FileSave"
    Case 15
        strImage = "SaveAll"
    Case 16
        strImage = "ObjectPictureFill"
    
    Case 21
        Set objImage = Get�C���[�W(Range("Resource!1:1"))
    Case 22
        strImage = "PasteAsPicture"
    Case 23
        strImage = "PasteAsEmbedded"
    Case 24
        strImage = "ObjectRotateRight90"
    Case 25
        strImage = "ObjectRotateLeft90"
    
    Case 31
        strImage = "BlackAndWhiteWhite"
    Case 32
        strImage = "ViewGridlines"
    Case 33
'        strImage = "FillStyle"
        Set objImage = Get�C���[�W(Range("Resource!8:8"))
    Case 34
        Set objImage = Get�C���[�W(Range("Resource!7:7"))
    Case 35
        strImage = "ObjectFlipHorizontal"
    Case 36
        strImage = "ObjectFlipVertical"
    Case 37
        strImage = "ObjectRotateRight90"
    Case 38
        strImage = "ObjectRotateLeft90"
    
    Case 41
'        strImage = "TableSelectCellInfoPath"
        Set objImage = Get�C���[�W(Range("Resource!2:2"))
    Case 42
        Set objImage = Get�C���[�W(Range("Resource!3:3"))
    Case 43
'        strImage = "SelectSheet"
        Set objImage = Get�C���[�W(Range("Resource!4:4"))
    
    Case 51
        Set objImage = Get�C���[�W(Range("Resource!9:9"))
    Case 52
        Set objImage = Get�C���[�W(Range("Resource!10:10"))
    Case 53
        Set objImage = Get�C���[�W(Range("Resource!11:11"))
    Case 54
        Set objImage = Get�C���[�W(Range("Resource!12:12"))
    Case 55
        Set objImage = Get�C���[�W(Range("Resource!13:13"))
    Case 56
        Set objImage = Get�C���[�W(Range("Resource!14:14"))
    Case 57
        strImage = "CatalogMergeGoToPreviousRecord"
    Case 58
        strImage = "CatalogMergeGoToNextRecord"
    
    Case 61 To 66
        If FSampleClick Then
            Set objImage = Get�T���v���摜()
            FSampleClick = False
        Else
            strImage = "TentativeAcceptInvitation"
        End If
    
    
    Case 74
        strImage = "EditFormula"
    Case 71, 72, 75
'        strImage = "ColorFuchsia"
        Set objImage = Get�C���[�W(Range("Resource!5:5"))
    Case 73, 76
'        strImage = "NotebookColor1"
        Set objImage = Get�C���[�W(Range("Resource!6:6"))
'    Case 77, 78
'        strImage = "DataTypeCalculatedColumn"
    Case Else
'        strimage = "BlackAndWhiteWhite"
    End Select

    If strImage = "" Then
        Set Result = objImage
    Else
        Result = strImage
    End If
End Sub

'*****************************************************************************
'[�C�x���g] onAction
'*****************************************************************************
Sub onAction(control As IRibbonControl)
'   Call GetRibbonUI.Invalidate
    
    Select Case Mid(control.ID, 2)
    Case 11
        Call ImageMso�擾
    Case 12
        Call �摜�Ǎ�
    Case 13
        Call �ꊇ���s�V�[�g���J��
    Case 14
        Call �摜�ۑ�
    Case 15
        Call �ꊇ���s�V�[�g���J��
    Case 16
        Call �ꊇ���s�V�[�g���J��
    
    Case 21
        Call �\�t��
    Case 22
        Call Clipbord�摜�Ǎ�
    Case 23
        Call Shape�Ǎ�
    Case 24
        Call ��](2, 90)
    Case 25
        Call ��](2, -90)
    
    Case 31
        Call �N���A
    Case 32
        Call �����F����
    Case 33
        Call �h�ׂ�
    Case 34
        Call �F�̒u��
    Case 35
        Call ���E���]
    Case 36
        Call �㉺���]
    Case 37
        Call ��](1, 90)
    Case 38
        Call ��](1, -90)
    
    Case 41
        Call ���F�I��(True)
    Case 42
        Call ���F�I��(False)
    Case 43
        Call �I�𔽓]��
    
    Case 51
        Call HSL����(-1, 3)
    Case 52
        Call HSL����(1, 3)
    Case 53
        Call HSL����(-1, 2)
    Case 54
        Call HSL����(1, 2)
    Case 55
        Call HSL����(-1, 1)
    Case 56
        Call HSL����(1, 1)
    Case 57
        Call �F����(-1)
    Case 58
        Call �F����(1)
    
    Case 61 To 66
        If CheckSelection <> E_Range Then Exit Sub
        Call Clipbord�摜�ݒ�
        FSampleClick = True
        Call GetRibbonUI.InvalidateControl(control.ID)

    Case 71
        Call �F�𐔒l��(False)
    Case 72
        Call �F�𐔒l��(True)
    Case 73
        Call �A���t�@�l��\��
    Case 74
        Application.ScreenUpdating = False
        Call Worksheets("���߂�").Activate
        ActiveWindow.ScrollColumn = 1
        ActiveWindow.ScrollRow = 90
        Call Worksheets("���߂�").Range("A90").Select
        Application.ScreenUpdating = True
    Case 75
        Call ���l����F��ݒ�
    Case 76
        Call ���l����A���t�@�l��ݒ�
    End Select
End Sub


'*****************************************************************************
'[�T�v] �Z���̃f�[�^����A�C�R���t�@�C����Ǎ���
'[����] �o�C�i���t�@�C���f�[�^���擾����s(Range�I�u�W�F�N�g)
'[�ߒl] IPicture
'*****************************************************************************
Private Function Get�C���[�W(ByRef objRange As Range) As IPicture
On Error GoTo ErrHandle
    Dim img As New CImage
    Call img.LoadImageFromResource(objRange)
    Set Get�C���[�W = img.SetToIPicture
ErrHandle:
End Function

'*****************************************************************************
'[�T�v] ���{���̃R�[���o�b�N�֐������s����(Debug�p)
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub InvalidateRibbon()
    Call GetRibbonUI.Invalidate
End Sub

