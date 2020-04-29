Attribute VB_Name = "Ribbon"
Option Explicit

Private FRibbon As IRibbonUI
Public FChecked(1 To 7) As Boolean
Private FSampleClick As Boolean

Sub onLoad(ribbon As IRibbonUI)
    Set FRibbon = ribbon
End Sub

Sub loadImage(imageID As String, ByRef returnedVal)
  returnedVal = imageID
End Sub

Sub getVisible(control As IRibbonControl, ByRef returnedVal)
    returnedVal = True
End Sub

Sub getEnabled(control As IRibbonControl, ByRef returnedVal)
    Select Case control.ID
    Case "C3"
        returnedVal = FChecked(1) Or FChecked(2)
    Case Else
        returnedVal = True
    End Select
End Sub

Sub getShowLabel(control As IRibbonControl, ByRef returnedVal)
    returnedVal = GetTips(control, 0)
End Sub

Sub getLabel(control As IRibbonControl, ByRef returnedVal)
    returnedVal = GetTips(control, 1)
End Sub

Sub getScreentip(control As IRibbonControl, ByRef returnedVal)
    returnedVal = GetTips(control, 2)
End Sub

Sub getSupertip(control As IRibbonControl, ByRef returnedVal)
    returnedVal = GetTips(control, 3)
End Sub

'*****************************************************************************
'[�T�v] Label�����ScreenTip��ݒ肵�܂�
'[����] lngType�u0:getShowLabel, 1:getLabel, 2:getScreentip, 3:getSupertip�v
'[�ߒl] �ݒ�l
'*****************************************************************************
Private Function GetTips(control As IRibbonControl, ByVal lngType As Long) As Variant
    ReDim Result(1 To 3) '1:getLabel, 2:getScreentip, 3:getSupertip
    Select Case Mid(control.ID, 2)
    Case 11
        Result(1) = "ImageMso"
        Result(2) = "ImageMso����摜���擾"
        Result(3) = "ImageMso���w�肵�đI�����ꂽ�Z���̈ʒu�։摜��ǂݍ��݂܂�"
    Case 12
        Result(1) = "�Ǎ�"
        Result(2) = "�摜�̓Ǎ�"
        Result(3) = "�I�����ꂽ�Z���̈ʒu�։摜��Ǎ��݂܂�"
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
        Result(3) = "�N���b�v�{�[�h�̉摜��\�t���܂�"
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
        Result(3) = "�I�����ꂽ�Z���͈̔͂�Ώۂɓh�ׂ��܂�"
    
    Case 61
        Result(1) = "�����F"
        Result(2) = "�����F�̃Z���̑I��"
        Result(3) = "�����F�̃Z����I�����܂�"
    Case 62
        Result(1) = "�Ⴄ�F"
        Result(2) = "�Ⴄ�F�̃Z���̑I��"
        Result(3) = "�Ⴄ�F�̃Z����I�����܂�"
    Case 63
        Result(1) = "���]��"
        Result(2) = "�I��̈�̔��]�Ȃ�"
        Result(3) = "�I��̈�̔��]��ꕔ���O�Ȃǂ��s���܂�"
    Case 71
        Result(1) = ""
        Result(2) = "RGB�e�F�̐��l����уA���t�@�l�����������܂�"
        Result(3) = "�I�������͈͂�RGB�e�F�̐��l(0�`255)��16(Ctrl��������1)���������܂�"
    Case 72
        Result(1) = ""
        Result(2) = "RGB�e�F�̐��l����уA���t�@�l�𑝉������܂�"
        Result(3) = "�I�������͈͂�RGB�e�F�̐��l(0�`255)��16(Ctrl��������1)���������܂�"
    
    Case 41 To 46
        Select Case Mid(control.ID, 2)
        Case 41, 44
            Result(1) = "��1"
        Case 42, 45
            Result(1) = "��2"
        Case 43, 46
            Result(1) = "��3"
        End Select
        Select Case Mid(control.ID, 2)
        Case 41 To 43
            Result(2) = "�T���v���\��(��)"
        Case Else
            Result(2) = "�T���v���\��(��)"
        End Select
        Result(3) = "�I�����ꂽ�Z���͈͂̉摜��\�����܂�" & vbLf & _
                    "���킹�ăN���b�v�{�[�h�ɂ��摜���R�s�[���܂�"
    Case 51
        Result(1) = "�����F�̋���"
        Result(2) = "�����F�̋���"
        Result(3) = "�I�����ꂽ�Z���͈͂̓���(������)�F�̃Z����Ԋ|���ŋ������܂�"
    Case 52
        Result(1) = "�F�����l"
        Result(2) = "�F�𐔒l��"
        Result(3) = "�F(RGB)��16�i���ŃZ���ɐݒ肵�܂�"
    Case 53
        Result(1) = "�F�����l(��)"
        Result(2) = "�F���A���t�@�`�����l���t���Ő��l��"
        Result(3) = "�F(ARGB)��16�i���ŃZ���ɐݒ肵�܂�"
    Case 54
        Result(1) = "���l���F"
        Result(2) = "�Z���̐��l����F��ݒ�"
        Result(3) = "�Z���̐��l��F�ɕϊ����܂�" & vbLf & _
                      "10�i���E16�i���̂�������Ή����Ă��܂�" & vbLf & _
                      "16�i����6���ȉ��̎��͂��ׂĕs�����F�ɐݒ肵�܂�" & vbLf & _
                      "10�i���̎��́A0�͓����F�ƍ��̔��ʂ����Ȃ��̂œ����F�Ƃ��܂�"
    Case 55
        Result(1) = "���l����"
        Result(2) = "�Z���̐��l����A���t�@�l��ݒ�"
        Result(3) = "�Z���̒l��(0�`255)�̎��A�������x(�A���t�@�l)��ݒ肵�܂�" & vbLf & _
                      "0�͊��S����" & vbLf & _
                      "255�͊��S�s�����ł�"
    Case 56
        Result(1) = "16�i��10�i"
        Result(2) = "16�i����10�i��"
        Result(3) = "�Z���̒l��16�i���̎��A10�i���ɕϊ����܂�"
    Case 57
        Result(1) = "10�i��16�i"
        Result(2) = "10�i����16�i��"
        Result(3) = "�Z���̒l��10�i���̎��A16�i���ɕϊ����܂�"
    Case Else
        Result(1) = ""
    End Select
    
    If lngType = 0 Then
        GetTips = (Result(1) <> "")
    Else
        GetTips = Result(lngType)
    End If
End Function

Sub getShowImage(control As IRibbonControl, ByRef returnedVal)
    Call GetImages(control, 0, returnedVal)
End Sub

Sub getImage(control As IRibbonControl, ByRef returnedVal)
    Call GetImages(control, 1, returnedVal)
End Sub

Sub getSize(control As IRibbonControl, ByRef returnedVal)
    Call GetImages(control, 2, returnedVal)
End Sub

'*****************************************************************************
'[�T�v] Image��ݒ肵�܂�
'[����] lngType�u0:getShowImage, 1:getImage, 2:getSize�v
'[�ߒl] Result
'*****************************************************************************
Private Sub GetImages(control As IRibbonControl, ByVal lngType As Long, ByRef Result)
    Dim strImage As String
    Dim lngSize  As Long
    Dim objImage As IPictureDisp

    Select Case Mid(control.ID, 2)
    Case 11
        strImage = "PictureReset"
        lngSize = 1 'large
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
        lngSize = 1 'large
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
    Case 36
'        strImage = "FillStyle"
        Set objImage = Get�C���[�W(Range("Icons!CW35:EB66"))
    Case 61
'        strImage = "TableSelectCellInfoPath"
        Set objImage = Get�C���[�W(Range("Icons!AI2:BN33"))
    Case 62
        Set objImage = Get�C���[�W(Range("Icons!BP2:CU33"))
    Case 63
'        strImage = "SelectSheet"
        Set objImage = Get�C���[�W(Range("Icons!CW2:EB33"))
    Case 71
        strImage = "CatalogMergeGoToPreviousRecord"
    Case 72
        strImage = "CatalogMergeGoToNextRecord"
    Case 41 To 46
        If FSampleClick Then
            Set objImage = Get�T���v���摜()
            FSampleClick = False
        Else
            strImage = "TentativeAcceptInvitation"
        End If
        If Mid(control.ID, 2) <= 43 Then
            lngSize = 1 'large
        Else
            lngSize = 0 'normal
        End If
    Case 51
        strImage = "ViewGridlines"
        lngSize = 1 'large
    Case 52
        strImage = "_1"
    Case 53
        strImage = "_2"
    Case 54
'        strImage = "ColorFuchsia"
        Set objImage = Get�C���[�W(Range("Icons!B35:AG66"))
    Case 55
'        strImage = "NotebookColor1"
        Set objImage = Get�C���[�W(Range("Icons!AI35:BN66"))
    Case 56, 57
        strImage = "DataTypeCalculatedColumn"
    Case Else
'        strimage = "BlackAndWhiteWhite"
    End Select

    Select Case lngType
    Case 0
        If (strImage = "") And (objImage Is Nothing) Then
            Result = False
        Else
            Result = True
        End If
    Case 1
        If strImage = "" Then
            Set Result = objImage
        Else
            Result = strImage
        End If
    Case 2
        Result = lngSize
    End Select
End Sub

Sub getPressed(control As IRibbonControl, ByRef returnedVal)
    Select Case Mid(control.ID, 2)
    Case 4 To 6
        returnedVal = True
        FChecked(Mid(control.ID, 2)) = True
    Case Else
        returnedVal = False
    End Select
End Sub

Sub onAction(control As IRibbonControl)
    Select Case Mid(control.ID, 2)
    Case 11
        Call ImageMso�擾
    Case 12
        Call �摜�Ǎ�
    Case 13
        Call FRibbon.Invalidate
    Case 14
        Call �摜�ۑ�
    Case 15
        Call FRibbon.Invalidate
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
    Case 61
        Call ���F�I��(True)
    Case 62
        Call ���F�I��(False)
    Case 63
        Call �I�𔽓]��
    Case 71
        Call �F����(-1)
    Case 72
        Call �F����(1)
    Case 41 To 46
        If CheckSelection <> E_Range Then Exit Sub
        Call Clipbord�摜�ݒ�
        FSampleClick = True
        Call FRibbon.InvalidateControl(control.ID)
    Case 51
        Call �����F����
    End Select
End Sub

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
        Call FRibbon.InvalidateControl("C2")
        
        '�L��������؂�ւ�
        Call FRibbon.InvalidateControl("C3")
'        Application.EnableEvents = True
    Case 2
'        Application.EnableEvents = False
        '����F�E����F�ȊO�̃g�O��
        FChecked(1) = False
        Call FRibbon.InvalidateControl("C1")
        
        '�L��������؂�ւ�
        Call FRibbon.InvalidateControl("C3")
'        Application.EnableEvents = True
    End Select
End Sub

Private Function Get�C���[�W(ByRef objRange As Range) As IPictureDisp
On Error GoTo ErrHandle
    Dim img As New CImage
    Call img.GetPixelsFromRange(objRange)
    Set Get�C���[�W = img.SetToIPicture
ErrHandle:
End Function
