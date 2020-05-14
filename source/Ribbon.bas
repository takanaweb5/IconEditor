Attribute VB_Name = "Ribbon"
Option Explicit

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
Public FChecked(1 To 7) As Boolean
Private FSampleClick As Boolean

'*****************************************************************************
'[�T�v] IRibbonUI�����L�������ɕۑ�����
'[����] IRibbonUI
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub SetRibbonUI(ByRef Ribbon As IRibbonUI)
    Dim hFileMap As LongPtr
    Dim pMap     As LongPtr
    Dim Pointer  As LongPtr

    '�n���h����Close���邱�Ƃ͂�����߂�
    hFileMap = CreateFileMapping(-1, ByVal 0&, PAGE_READWRITE, 0, Len(Pointer), ThisWorkbook.FullName)
'    hFileMap = OpenFileMapping(FILE_MAP_WRITE, False, ThisWorkbook.FullName)
    If hFileMap <> 0 Then
        pMap = MapViewOfFile(hFileMap, FILE_MAP_WRITE, 0, 0, 0)
        If pMap <> 0 Then
            Pointer = ObjPtr(Ribbon)
            Call CopyMemory(ByVal pMap, Pointer, Len(Pointer))
            Call UnmapViewOfFile(pMap)
        End If
    End If
'    Set FRibbon = Ribbon
End Sub

'*****************************************************************************
'[�T�v] IRibbonUI�����L����������擾����
'[����] �Ȃ�
'[�ߒl] IRibbonUI
'*****************************************************************************
Private Function GetRibbonUI() As IRibbonUI
    Dim hFileMap As LongPtr
    Dim pMap     As LongPtr
    Dim Pointer  As LongPtr

    hFileMap = OpenFileMapping(FILE_MAP_READ, False, ThisWorkbook.FullName)
    If hFileMap <> 0 Then
        pMap = MapViewOfFile(hFileMap, FILE_MAP_READ, 0, 0, 0)
        If pMap <> 0 Then
            Call CopyMemory(Pointer, ByVal pMap, Len(Pointer))
            Call UnmapViewOfFile(pMap)

            Dim obj As Object
            Call CopyMemory(obj, Pointer, Len(Pointer))
            Set GetRibbonUI = obj
        End If
        Call CloseHandle(hFileMap)
    End If
'    GetRibbonUI = FRibbon
End Function

'*****************************************************************************
'[�C�x���g] onLoad
'*****************************************************************************
Sub onLoad(Ribbon As IRibbonUI)
    '���{��UI�����L�������ɕۑ�����
    '(���W���[���ϐ��ɕۑ������ꍇ�́A��O��R�[�h��Break�Œl�����Ȃ��邽��)
    Call SetRibbonUI(Ribbon)
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
        returnedVal = FChecked(1) Or FChecked(2)
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
    Case 11, 21, 37, 61, 62, 63, 71
        returnedVal = 1
    Case Else
        returnedVal = 0
    End Select
End Sub

'*****************************************************************************
'[�C�x���g] getPressed
'*****************************************************************************
Sub getPressed(control As IRibbonControl, ByRef returnedVal)
    Select Case Mid(control.ID, 2)
    Case 4 To 6
        returnedVal = True
        FChecked(Mid(control.ID, 2)) = True
    Case Else
        returnedVal = False
    End Select
End Sub

'*****************************************************************************
'[�C�x���g] onCheckAction
'*****************************************************************************
Sub onCheckAction(control As IRibbonControl, pressed As Boolean)
    Dim ID As Long
    ID = Mid(control.ID, 2)
    
    '�`�F�b�N��Ԃ�ۑ�
    FChecked(ID) = pressed
    Select Case ID
    Case 1
'        Application.EnableEvents = False
        '����F�E����F�ȊO�̃g�O��
        FChecked(2) = False
        Call GetRibbonUI.InvalidateControl("C2")
        
        '�L��������؂�ւ�
        Call GetRibbonUI.InvalidateControl("C3")
'        Application.EnableEvents = True
    Case 2
'        Application.EnableEvents = False
        '����F�E����F�ȊO�̃g�O��
        FChecked(1) = False
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
        Result(1) = "�\�t"
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
        Result(1) = ""
        Result(2) = "���E���]"
        Result(3) = "�I�����ꂽ�Z���͈͂����E���]���܂�"
    Case 32
        Result(1) = ""
        Result(2) = "�㉺���]"
        Result(3) = "�I�����ꂽ�Z���͈͂��㉺���]���܂�"
    Case 33
        Result(1) = ""
        Result(2) = "�E��90����]"
        Result(3) = "�I�����ꂽ�Z���͈͂��E��90����]���܂�"
    Case 34
        Result(1) = ""
        Result(2) = "����90����]"
        Result(3) = "�I�����ꂽ�Z���͈͂�����90����]���܂�"
    Case 35
        Result(1) = "�F�̒u��"
        Result(2) = "�F�̒u��"
        Result(3) = "�I�����ꂽ�Z���͈͂̐F��u�����܂�"
    Case 36
        Result(1) = "�h�ׂ�"
        Result(2) = "�h�ׂ�"
        Result(3) = "�I�����ꂽ�Z���͈͂�Ώۂɓh�ׂ��܂�"
    Case 37
        Result(1) = "�N���A"
        Result(2) = "�N���A"
        Result(3) = "�I�����ꂽ�Z���͈͂̐F���N���A���܂�"
    
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
        Result(1) = "�����F�̋���"
        Result(2) = "�����F�̋���"
        Result(3) = "�I�����ꂽ�Z���͈͂̓���(������)�F�̃Z����Ԋ|���ŋ������܂�"
    Case 72
        Result(1) = "�F�����l"
        Result(2) = "�F�𐔒l��"
        Result(3) = "�F(RGB)��16�i���ŃZ���ɐݒ肵�܂�"
    Case 73
        Result(1) = "�F�����l(��)"
        Result(2) = "�F���A���t�@�`�����l���t���Ő��l��"
        Result(3) = "�F(ARGB)��16�i���ŃZ���ɐݒ肵�܂�"
    Case 74
        Result(1) = "���l���F"
        Result(2) = "�Z���̐��l����F��ݒ�"
        Result(3) = "�Z���̐��l��F�ɕϊ����܂�" & vbLf & _
                      "10�i���E16�i���̂�������Ή����Ă��܂�" & vbLf & _
                      "16�i����6���ȉ��̎��͂��ׂĕs�����F�ɐݒ肵�܂�" & vbLf & _
                      "10�i���̎��́A0�͓����F�ƍ��̔��ʂ����Ȃ��̂œ����F�Ƃ��܂�"
    Case 75
        Result(1) = "���l����"
        Result(2) = "�Z���̐��l����A���t�@�l��ݒ�"
        Result(3) = "�Z���̒l��(0�`255)�̎��A�������x(�A���t�@�l)��ݒ肵�܂�" & vbLf & _
                      "0�͊��S����" & vbLf & _
                      "255�͊��S�s�����ł�"
    Case 76
        Result(1) = "16�i��10�i"
        Result(2) = "16�i����10�i��"
        Result(3) = "�Z���̒l��16�i���̎��A10�i���ɕϊ����܂�"
    Case 77
        Result(1) = "10�i��16�i"
        Result(2) = "10�i����16�i��"
        Result(3) = "�Z���̒l��10�i���̎��A16�i���ɕϊ����܂�"
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
        Set objImage = Get�C���[�W(Range("Icons!B2:AG33"))
    Case 22
        strImage = "PasteAsPicture"
    Case 23
        strImage = "PasteAsEmbedded"
    Case 24
        strImage = "ObjectRotateRight90"
    Case 25
        strImage = "ObjectRotateLeft90"
    Case 31
        strImage = "ObjectFlipHorizontal"
    Case 32
        strImage = "ObjectFlipVertical"
    Case 33
        strImage = "ObjectRotateRight90"
    Case 34
        strImage = "ObjectRotateLeft90"
    Case 35
        Set objImage = Get�C���[�W(Range("Icons!BP35:CU66"))
    Case 37
        strImage = "ViewGridlines"
    Case 36
'        strImage = "FillStyle"
        Set objImage = Get�C���[�W(Range("Icons!CW35:EB66"))
    Case 41
'        strImage = "TableSelectCellInfoPath"
        Set objImage = Get�C���[�W(Range("Icons!AI2:BN33"))
    Case 42
        Set objImage = Get�C���[�W(Range("Icons!BP2:CU33"))
    Case 43
'        strImage = "SelectSheet"
        Set objImage = Get�C���[�W(Range("Icons!CW2:EB33"))
    
    Case 51
        Set objImage = Get�C���[�W(Range("Icons!B68:AG99"))
    Case 52
        Set objImage = Get�C���[�W(Range("Icons!AI68:BN99"))
    
    Case 53
        Set objImage = Get�C���[�W(Range("Icons!B101:AG132"))
    Case 54
        Set objImage = Get�C���[�W(Range("Icons!AI101:BN132"))
    Case 55
        Set objImage = Get�C���[�W(Range("Icons!B134:AG165"))
    Case 56
        Set objImage = Get�C���[�W(Range("Icons!AI134:BN165"))
    
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
    Case 71
        strImage = "ViewGridlines"
    Case 72
        strImage = "_1"
    Case 73
        strImage = "_2"
    Case 74
'        strImage = "ColorFuchsia"
        Set objImage = Get�C���[�W(Range("Icons!B35:AG66"))
    Case 75
'        strImage = "NotebookColor1"
        Set objImage = Get�C���[�W(Range("Icons!AI35:BN66"))
    Case 76, 77
        strImage = "DataTypeCalculatedColumn"
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
    Select Case Mid(control.ID, 2)
    Case 11
        Call ImageMso�擾
    Case 12
        Call �摜�Ǎ�
    Case 13
        Call GetRibbonUI.Invalidate
    Case 14
        Call �摜�ۑ�
    Case 15
        Call GetRibbonUI.Invalidate
    Case 16
        Call Clipbord�摜�ۑ�
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
        Call ���E���]
    Case 32
        Call �㉺���]
    Case 33
        Call ��](1, 90)
    Case 34
        Call ��](1, -90)
    Case 35
        Call �F�̒u��
    Case 36
        Call �h�ׂ�
    Case 37
        Call �N���A
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
        Call �����F����
    End Select
End Sub

'*****************************************************************************
'[�C�x���g] loadImage
'*****************************************************************************
Private Function Get�C���[�W(ByRef objRange As Range) As IPicture
On Error GoTo ErrHandle
    Dim img As New CImage
    Call img.GetPixelsFromRange(objRange)
    Set Get�C���[�W = img.SetToIPicture
ErrHandle:
End Function
