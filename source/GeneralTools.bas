Attribute VB_Name = "GeneralTools"
Option Explicit
Option Private Module

Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As Long
Public Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long
Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Public Declare PtrSafe Function IsZoomed Lib "user32" (ByVal hwnd As LongPtr) As Long
Public Declare PtrSafe Function GetSystemMenu Lib "user32" (ByVal hwnd As LongPtr, ByVal bRevert As Long) As LongPtr
Public Declare PtrSafe Function EnableMenuItem Lib "user32.dll" (ByVal hMenu As LongPtr, ByVal uIDEnableItem As Long, ByVal uEnable As Long) As Long
Public Declare PtrSafe Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As LongPtr
Public Declare PtrSafe Function GetExitCodeProcess Lib "kernel32.dll" (ByVal hProcess As LongPtr, lpExitCode As Long) As Long
Public Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
Public Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As Long
Public Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hDC As LongPtr) As Long

Public Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
Public Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Public Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
Public Declare PtrSafe Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
Public Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Public Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As Long
Public Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As LongPtr
Public Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As LongPtr
Public Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As Long
Public Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As Long
Public Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Public Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Public Declare PtrSafe Function GetKeyState Lib "user32" (ByVal lngVirtKey As Long) As Integer
Public Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long

Public Declare PtrSafe Function ImmGetContext Lib "imm32.dll" (ByVal hwnd As LongPtr) As LongPtr
Public Declare PtrSafe Function ImmSetOpenStatus Lib "imm32.dll" (ByVal himc As LongPtr, ByVal B As Long) As Long
Public Declare PtrSafe Function ImmReleaseContext Lib "imm32.dll" (ByVal hwnd As LongPtr, ByVal himc As LongPtr) As Long

Public Declare PtrSafe Function CopyEnhMetaFile Lib "gdi32" Alias "CopyEnhMetaFileA" (ByVal hemfSrc As LongPtr, ByVal lpszFile As String) As LongPtr
Public Declare PtrSafe Function DeleteEnhMetaFile Lib "gdi32" (ByVal hEmf As LongPtr) As Long

Public Const CF_BITMAP = 2
Public Const CF_ENHMETAFILE = 14

'�I���^�C�v
Public Enum ESelectionType
    E_Range
    E_Shape
    E_Non
    E_Other
End Enum

Public Const MAX_WIDTH = 256
Public Const MAX_HEIGHT = 256

' �萔�̒�`
Public Const SC_CLOSE = 61536
Public Const SC_SIZE = &HF000&
Public Const MF_BYCOMMAND = 0
Public Const MF_GRAYED = 1
Public Const LOGPIXELSX = 88
Public Const LOGPIXELSY = 90

'�\�[�g�p�\����
Public Type TSortArray
    Key1  As Long
    Key2  As Long
    Key3  As Long
End Type

'*****************************************************************************
'[�T�v] �I������Ă��邩�I�u�W�F�N�g�̎�ނ𔻒肷��
'[����] �Ȃ�
'[�ߒl] Range�AShape�A���̑��@�̂����ꂩ
'*****************************************************************************
Public Function CheckSelection() As ESelectionType
On Error GoTo ErrHandle
    If ActiveWorkbook Is Nothing Then
        CheckSelection = E_Non
        Exit Function
    End If
    
    If Selection Is Nothing Then
        CheckSelection = E_Other
        Exit Function
    End If
    
    If TypeOf Selection Is Range Then
        CheckSelection = E_Range
    ElseIf TypeOf Selection.ShapeRange Is ShapeRange Then
        CheckSelection = E_Shape
    Else
        CheckSelection = E_Other
    End If
Exit Function
ErrHandle:
    CheckSelection = E_Other
End Function

'*****************************************************************************
'[�T�v] �R�s�[�Ώۂ�Range���擾����
'[����] �Ȃ�
'[�ߒl] �R�s�[�Ώۂ�Range
'*****************************************************************************
Public Function GetCopyRange() As Range
    If OpenClipboard(0) = 0 Then Exit Function
    Dim hMem As LongPtr
    hMem = GetClipboardData(RegisterClipboardFormat("Link"))
    If hMem = 0 Then
        Call CloseClipboard
        Exit Function
    End If
     
On Error GoTo ErrHandle
    Dim size As Long
    Dim p As LongPtr
    size = GlobalSize(hMem)
    p = GlobalLock(hMem)
    ReDim Data(1 To size) As Byte
    Call CopyMemory(Data(1), ByVal p, size)
    Call GlobalUnlock(hMem)
    Call CloseClipboard
    hMem = 0
    
    Dim strData As String
    Dim i As Long
    For i = 1 To size
        If Data(i) = 0 Then
            Data(i) = Asc("/") '�V�[�g���ɂ��t�@�C�����ɂ��g���Ȃ�����
        End If
    Next i
    strData = StrConv(Data, vbUnicode)
    
    Dim objRegExp As Object
    Set objRegExp = CreateObject("VBScript.RegExp")
    objRegExp.Global = False
    objRegExp.Pattern = "Excel\/.*\[(.+)\](.+)\/(.+)\/\/$"
    If Not objRegExp.Test(strData) Then Exit Function
    With objRegExp.Execute(strData)(0)
        Dim strRange As String
        strRange = Application.ConvertFormula(.SubMatches(2), xlR1C1, xlA1)
        Set GetCopyRange = Workbooks(.SubMatches(0)).Worksheets(.SubMatches(1)).Range(strRange)
    End With
    Application.CutCopyMode = False
    Exit Function
ErrHandle:
    If hMem <> 0 Then Call CloseClipboard
End Function

'*****************************************************************************
'[�T�v] �N���b�v�{�[�h��Bitmap�`�����R�s�[����Ă��邩�ǂ���
'[����] �Ȃ�
'[�ߒl] True:Bitmap�`������
'*****************************************************************************
Public Function ClipboardHasBitmap() As Boolean
    ClipboardHasBitmap = (IsClipboardFormatAvailable(CF_BITMAP) <> 0)
End Function

'*****************************************************************************
'[�T�v] �t�H�[����\�����ăZ����I��������
'[����] �\�����郁�b�Z�[�W�AobjCurrentCell�F�����I��������Z��
'[�ߒl] �I�����ꂽ�Z���i�L�����Z������Nothing�j
'*****************************************************************************
Public Function SelectCell(ByVal strMsg As String, ByRef objCurrentCell As Range) As Range
    Dim strCell As String
    '�t�H�[����\��
    With frmSelectCell
        .Label.Caption = strMsg
        Call objCurrentCell.Worksheet.Activate
        .RefEdit.Text = objCurrentCell.AddressLocal
        Call .Show
        If .IsOK = True Then
            strCell = .RefEdit
        End If
    End With
    Call Unload(frmSelectCell)
    If strCell <> "" Then
        Set SelectCell = Range(strCell)
        If SelectCell.Address = SelectCell.Cells(1, 1).MergeArea.Address Then
            Set SelectCell = SelectCell.Cells(1, 1)
        End If
    End If
End Function

'*****************************************************************************
'[�T�v] �g���q�̎擾
'[����] �t�@�C���p�X
'[�ߒl] �g���q(�啶��)
'*****************************************************************************
Public Function GetFileExtension(ByVal strFilename As String) As String
    With CreateObject("Scripting.FileSystemObject")
        GetFileExtension = UCase(.GetExtensionName(strFilename))
    End With
End Function

'*****************************************************************************
'[�T�v] �̈�Ɨ̈�̏d�Ȃ�̈���擾����
'[����] �Ώۗ̈�(Nothing����)
'[�ߒl] objRange1 �� objRange2
'*****************************************************************************
Public Function IntersectRange(ByRef objRange1 As Range, ByRef objRange2 As Range) As Range
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
Public Function UnionRange(ByRef objRange1 As Range, ByRef objRange2 As Range) As Range
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
Public Function MinusRange(ByRef objRange1 As Range, ByRef objRange2 As Range) As Range
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
Public Function ReSelectRange(ByRef objRange As Range) As Range
    Set ReSelectRange = objRange.Areas(1)
    
    Dim i As Long
    For i = 2 To objRange.Areas.Count
        Set ReSelectRange = Union(ReSelectRange, ReSelectRange(MinusRange(objRange.Areas(i), ReSelectRange)))
    Next
End Function

'*****************************************************************************
'[�T�v] �̈悪��v���邩����
'[����] �Ώۗ̈�A�h���X
'[�ߒl] True:��v
'*****************************************************************************
'Public Function IsSameRange(ByRef strRange1 As String, ByRef strRange2 As String) As Boolean
'    If strRange1 = "" Or strRange2 = "" Then
'        Exit Function
'    End If
'
'    Dim objRange1 As Range
'    Dim objRange2 As Range
'    Set objRange1 = AddressToRange(strRange1)
'    Set objRange2 = AddressToRange(strRange2)
'    IsSameRange = MinusRange(objRange1, objRange2) Is Nothing
'    If IsSameRange Then
'        IsSameRange = MinusRange(objRange2, objRange1) Is Nothing
'    End If
'End Function

'*****************************************************************************
'[�T�v] Range�̃A�h���X���擾����(255���ȏ�ɑΉ����邽��)
'[����] Range
'[�ߒl] ��FA1:C3/E5/F1:G5
'*****************************************************************************
Public Function RangeToAddress(ByRef objRange As Range) As String
    ReDim Address(1 To objRange.Areas.Count)
    Dim i As Long
    For i = 1 To objRange.Areas.Count
        Address(i) = objRange.Areas(i).Address(False, False)
    Next
    RangeToAddress = Join(Address, "/")
End Function

'*****************************************************************************
'[�T�v] RangeToAddress�̌��ʂ���Range���擾����
'[����] ��FA1:C3/E5/F1:G5
'[�ߒl] Range
'*****************************************************************************
Public Function AddressToRange(ByVal strAddress As String) As Range
    Dim Address As Variant
    Address = Split(strAddress, "/")
    Dim i As Long
    For i = LBound(Address) To UBound(Address)
        Set AddressToRange = UnionRange(AddressToRange, Range(Address(i)))
    Next
End Function

'*****************************************************************************
'[�T�v] �Z���̐F���N���A����
'[����] �Ώۗ̈�
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Function ClearRange(ByRef objRange As Range)
    If objRange Is Nothing Then Exit Function
    With objRange
        .Interior.Pattern = xlNone
        .Font.ColorIndex = xlAutomatic
        .ClearContents
    End With
End Function

'*****************************************************************************
'[�T�v] �e���|������CommandBarControl���擾����
'[����] Control�����ʂ���ID�i���{���R���g���[����ID�j
'[�ߒl] CommandBarControl
'*****************************************************************************
Public Function GetTmpControl(ByVal strID As String) As CommandBarControl
    Set GetTmpControl = CommandBars.FindControl(, , strID & ThisWorkbook.Name)
End Function

'*****************************************************************************
'[�T�v] �o�C�i���t�@�C�����Z���ɓǍ���
'[����] �Ǎ��ރt�@�C����, �o�C�i���t�@�C����Ǎ��ލs(Range�I�u�W�F�N�g)
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub LoadResourceFromFile(ByVal strFilename As String, ByRef objRow As Range)
    'A��̓t�@�C�����Ƃ���
    objRow.Cells(1, 1).Value = Dir(strFilename)
    
    '�t�@�C���T�C�Y�̔z����쐬
    ReDim Data(1 To FileLen(strFilename)) As Byte

    Dim File As Integer
    File = FreeFile()
    Open strFilename For Binary Access Read As #File
    Get #File, , Data
    Close #File
    
    Dim x As Long
    For x = 1 To UBound(Data)
        objRow.Cells(1, x + 1) = Data(x)
    Next
End Sub

'*****************************************************************************
'[�T�v] �Z���̃f�[�^���o�C�i���t�@�C����������
'[����] �����ރt�@�C����, �o�C�i���t�@�C���f�[�^���擾����s(Range�I�u�W�F�N�g)
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub SaveResourceToFile(ByVal strFilename As String, ByRef objRow As Range)
    '�t�@�C���T�C�Y�̔z����쐬
    ReDim Data(1 To objRow.Cells(1, 1).End(xlToRight).Column - 1) As Byte
    Dim x As Long
    For x = 1 To UBound(Data)
         Data(x) = objRow.Cells(1, x + 1)
    Next
    
    Dim File As Integer
    File = FreeFile()
    Open strFilename For Binary Access Write As #File
    Put #File, , Data
    Close #File
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
                GetUndoStr = Trim(.List(1))
            End If
        End If
    End With
End Function

'*****************************************************************************
'[�T�v] �ύX�Ώۂ̐}�`�̒��ŉ�]���Ă�����̂��O���[�v������
'[����] �O���[�v���O�̐}�`
'[�ߒl] �O���[�v����̐}�`
'*****************************************************************************
Public Function GroupSelection(ByRef objShapes As ShapeRange) As ShapeRange
    Dim i            As Long
    Dim objShape     As Shape
    Dim btePlacement As Byte
    ReDim blnRotation(1 To objShapes.Count) As Boolean
    ReDim lngIDArray(1 To objShapes.Count) As Variant
    
    '�}�`�̐��������[�v
    For i = 1 To objShapes.Count
        Set objShape = objShapes(i)
        lngIDArray(i) = objShape.ID
        
        Select Case objShape.Rotation
        Case 90, 270, 180
            blnRotation(i) = True
        End Select
    Next

    '�}�`�̐��������[�v
    For i = 1 To objShapes.Count
        If blnRotation(i) = True Then
            Set objShape = GetShapeFromID(lngIDArray(i))
            btePlacement = objShape.Placement
            '�T�C�Y�ƈʒu������̃N���[�����쐬���O���[�v������
            With objShape.Duplicate
                .Top = objShape.Top
                .Left = objShape.Left
                If objShape.Top < 0 Then
                    '�}�`����]���č��W���}�C�i�X�ɂȂ������[���ɂȂ邽�ߕ␳����
                    Call .IncrementTop(objShape.Top)
                End If
                If objShape.Left < 0 Then
                    '�}�`����]���č��W���}�C�i�X�ɂȂ������[���ɂȂ邽�ߕ␳����
                    Call .IncrementLeft(objShape.Left)
                End If
                
                '�����ɂ���
                .Fill.Visible = msoFalse
                .Line.Visible = msoFalse
                With GetShapeRangeFromID(Array(.ID, objShape.ID)).Group
                    .AlternativeText = "EL_TemporaryGroup" & i
                    .Placement = btePlacement
                    lngIDArray(i) = .ID
                End With
            End With
        End If
    Next
    
    Set GroupSelection = GetShapeRangeFromID(lngIDArray)
End Function

'*****************************************************************************
'[�T�v] �ύX�Ώۂ̐}�`�̒��ŃO���[�v���������̂����ɖ߂�
'[����] �O���[�v�����O�̐}�`
'[�ߒl] �O���[�v������̐}�`
'*****************************************************************************
Public Function UnGroupSelection(ByRef objGroups As ShapeRange) As ShapeRange
    Dim i            As Long
    Dim btePlacement As Byte
    Dim objShape     As Shape
    ReDim blnRotation(1 To objGroups.Count) As Boolean
    ReDim lngIDArray(1 To objGroups.Count) As Variant
    
    '�}�`�̐��������[�v
    For i = 1 To objGroups.Count
        Set objShape = objGroups(i)
        lngIDArray(i) = objShape.ID
        
        If Left$(objShape.AlternativeText, 17) = "EL_TemporaryGroup" Then
            blnRotation(i) = True
        End If
    Next

    '�}�`�̐��������[�v
    For i = 1 To objGroups.Count
        If blnRotation(i) = True Then
            Set objShape = GetShapeFromID(lngIDArray(i))
            btePlacement = objShape.Placement
            With objShape.Ungroup
                .Item(1).Placement = btePlacement
                Call .Item(2).Delete
                lngIDArray(i) = .Item(1).ID
            End With
        End If
    Next i
    
    Set UnGroupSelection = GetShapeRangeFromID(lngIDArray)
End Function

'*****************************************************************************
'[�T�v] Shape�I�u�W�F�N�g��ID����Shape�I�u�W�F�N�g���擾
'[����] ID
'[�ߒl] Shape�I�u�W�F�N�g
'*****************************************************************************
Private Function GetShapeFromID(ByVal lngID As Long) As Shape
    Dim j As Long
    Dim lngIndex As Long
        
    For j = 1 To ActiveSheet.Shapes.Count
        If ActiveSheet.Shapes(j).ID = lngID Then
            lngIndex = j
            Exit For
        End If
    Next j
    
    Set GetShapeFromID = ActiveSheet.Shapes.Range(j).Item(1)
End Function

'*****************************************************************************
'[�T�v] Shpes�I�u�W�F�N�g��ID����ShapeRange�I�u�W�F�N�g���擾
'[����] ID�̔z��
'[�ߒl] ShapeRange�I�u�W�F�N�g
'*****************************************************************************
Public Function GetShapeRangeFromID(ByRef lngID As Variant) As ShapeRange
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
'[�T�v] Shape�̎l���ɍł��߂��Z���͈͂��擾����
'[����] Shape�I�u�W�F�N�g
'[�ߒl] �Z���͈�
'*****************************************************************************
Public Function GetNearlyRange(ByRef objShape As Shape) As Range
    Dim objTopLeft     As Range
    Dim objBottomRight As Range
    Set objTopLeft = objShape.TopLeftCell
    Set objBottomRight = objShape.BottomRightCell
    
    '��̈ʒu�ƍ�����ݒ�
    If objShape.Height = 0 Then
        With objTopLeft
            If .Top + .Height / 2 < objShape.Top Then
                Set objTopLeft = Cells(.Row + 1, .Column)
                Set objBottomRight = Cells(.Row + 1, objBottomRight.Column)
            End If
        End With
    Else
        '���̃Z���̍Đݒ�
        With objBottomRight
            If .Top = objShape.Top + objShape.Height Then
                Set objBottomRight = Cells(.Row - 1, .Column)
            End If
        End With
            
        '��[�̍Đݒ�
        With objTopLeft
            If .Top + .Height / 2 < objShape.Top Then
                If .Row + 1 <= objBottomRight.Row Then
                    Set objTopLeft = Cells(.Row + 1, .Column)
                End If
            End If
        End With
                
        '���[�̍Đݒ�
        With objBottomRight
            If .Top + .Height / 2 > objShape.Top + objShape.Height Then
                If .Row - 1 >= objTopLeft.Row Then
                    Set objBottomRight = Cells(.Row - 1, .Column)
                End If
            End If
        End With
    End If
    
    '���̈ʒu�ƕ���ݒ�
    If objShape.Width = 0 Then
        With objTopLeft
            If .Left + .Width / 2 < objShape.Left Then
                Set objTopLeft = Cells(.Row, .Column + 1)
                Set objBottomRight = Cells(objBottomRight.Row, .Column + 1)
            End If
        End With
    Else
        '�E�̃Z���̍Đݒ�
        With objBottomRight
            If .Left = objShape.Left + objShape.Width Then
                Set objBottomRight = Cells(.Row, .Column - 1)
            End If
        End With
    
        '���[�̍Đݒ�
        With objTopLeft
            If .Left + .Width / 2 < objShape.Left Then
                If .Column + 1 <= objBottomRight.Column Then
                    Set objTopLeft = Cells(.Row, .Column + 1)
                End If
            End If
        End With
                
        '�E�[�̍Đݒ�
        With objBottomRight
            If .Left + .Width / 2 > objShape.Left + objShape.Width Then
                If .Column - 1 >= objTopLeft.Column Then
                    Set objBottomRight = Cells(.Row, .Column - 1)
                End If
            End If
        End With
    End If
    
    Set GetNearlyRange = Range(objTopLeft, objBottomRight)
End Function

'*****************************************************************************
'[�T�v] DPI�̕ϊ������擾���� ��72(Excel�̃f�t�H���g��DPI)/��ʂ�DPI
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Function DPIRatio() As Double
    DPIRatio = 72 / GetDPI()
End Function

'*****************************************************************************
'[�T�v] DPI���擾����
'[����] �Ȃ�
'[�ߒl] DPI ���W����96
'*****************************************************************************
Public Function GetDPI() As Long
    Dim DC As LongPtr
    DC = GetDC(0)
    GetDPI = GetDeviceCaps(DC, LOGPIXELSX)
    Call ReleaseDC(0, DC)
End Function

'*****************************************************************************
'[�T�v] SortArray�z����\�[�g����
'[����] Sort�Ώ۔z��
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub SortArray(ByRef SortArray() As TSortArray)
    '�o�u���\�[�g
    Dim i As Long
    Dim j As Long
    Dim Swap As TSortArray
    For i = UBound(SortArray) To 1 Step -1
        For j = 1 To i - 1
            If CompareValue(SortArray(j), SortArray(j + 1)) Then
                Swap = SortArray(j)
                SortArray(j) = SortArray(j + 1)
                SortArray(j + 1) = Swap
            End If
        Next j
    Next i
End Sub

'*****************************************************************************
'[�T�v] �召��r���s��
'[����] ��r�Ώ�
'[�ߒl] True: SortArray1 > SortArray2
'*****************************************************************************
Private Function CompareValue(ByRef SortArray1 As TSortArray, ByRef SortArray2 As TSortArray) As Boolean
    If SortArray1.Key1 = SortArray2.Key1 Then
        CompareValue = (SortArray1.Key2 > SortArray2.Key2)
    Else
        CompareValue = (SortArray1.Key1 > SortArray2.Key1)
    End If
End Function

'*****************************************************************************
'[�T�v] ���͂̈ʒu�̍����̘g���̈ʒu���擾(�P�ʃs�N�Z��)
'[����] lngPos:�ʒu(�P�ʃs�N�Z��)
'       objColumn: lngPos���܂ޗ�
'[�ߒl] �}�`�̍����̘g���̈ʒu(�P�ʃs�N�Z��)
'*****************************************************************************
Public Function GetLeftGrid(ByVal lngPos As Long, ByRef objColumn As Range) As Long
    Dim i       As Long
    Dim lngLeft As Long
    
    If lngPos <= Round(Columns(2).Left / DPIRatio) Then
        GetLeftGrid = 0
        Exit Function
    End If
    
    For i = objColumn.Column To 1 Step -1
        lngLeft = Round(GetWidth(Range(Columns(1), Columns(i - 1))) / DPIRatio)
        If lngLeft < lngPos Then
            GetLeftGrid = lngLeft
            Exit Function
        End If
    Next
End Function

'*****************************************************************************
'[�T�v] ���͂̈ʒu�̉E���̘g���̈ʒu���擾(�P�ʃs�N�Z��)
'[����] lngPos:�ʒu(�P�ʃs�N�Z��)
'       objColumn: lngPos���܂ޗ�
'[�ߒl] �}�`�̉E���̘g���̈ʒu(�P�ʃs�N�Z��)
'*****************************************************************************
Public Function GetRightGrid(ByVal lngPos As Long, ByRef objColumn As Range) As Long
    Dim i        As Long
    Dim lngRight As Long
    
    If lngPos >= Round(GetWidth(Range(Columns(1), Columns(Columns.Count - 1))) / DPIRatio) Then
        GetRightGrid = Round(GetWidth(Columns) / DPIRatio)
        Exit Function
    End If
    
    For i = objColumn.Column + 1 To Columns.Count
        lngRight = Round(GetWidth(Range(Columns(1), Columns(i - 1))) / DPIRatio)
        If lngRight > lngPos Then
            GetRightGrid = lngRight
            Exit Function
        End If
    Next
End Function

'*****************************************************************************
'[�T�v] ���͂̈ʒu�̏�̘g���̈ʒu���擾(�P�ʃs�N�Z��)
'[����] lngPos:�ʒu(�P�ʃs�N�Z��)
'       objRow: lngPos���܂ލs
'[�ߒl] �}�`�̏㑤�̘g���̈ʒu(�P�ʃs�N�Z��)
'*****************************************************************************
Public Function GetTopGrid(ByVal lngPos As Long, ByRef objRow As Range) As Long
    Dim i      As Long
    Dim lngTop As Long
    
    If lngPos <= Round(Rows(2).Top / DPIRatio) Then
        GetTopGrid = 0
        Exit Function
    End If
    
    For i = objRow.Row To 1 Step -1
        lngTop = Round(Rows(i).Top / DPIRatio)
        If lngTop < lngPos Then
            GetTopGrid = lngTop
            Exit Function
        End If
    Next
End Function

'*****************************************************************************
'[�T�v] ���͂̈ʒu�̉��̘g���̈ʒu���擾(�P�ʃs�N�Z��)
'[����] lngPos:�ʒu(�P�ʃs�N�Z��)
'       objRow: lngPos���܂ލs
'[�ߒl] �}�`�̉����̘g���̈ʒu(�P�ʃs�N�Z��)
'*****************************************************************************
Public Function GetBottomGrid(ByVal lngPos As Long, ByRef objRow As Range) As Long
    Dim i         As Long
    Dim lngBottom As Long
    Dim lngMax    As Long
    
    lngMax = Round((Rows(Rows.Count).Top + Rows(Rows.Count).Height) / DPIRatio)
    
    If lngPos >= Round(Rows(Rows.Count).Top / DPIRatio) Then
        GetBottomGrid = lngMax
        Exit Function
    End If
    
    For i = objRow.Row + 1 To Rows.Count
        lngBottom = Round(Rows(i).Top / DPIRatio)
        If lngBottom > lngPos Then
            GetBottomGrid = lngBottom
            Exit Function
        End If
    Next
End Function

'*****************************************************************************
'[�T�v] �I���G���A�̕����擾
'       Width/Left�v���p�e�B��32767�ȏ�̕����v�Z�o���Ȃ�����
'[����] �����擾����G���A
'[�ߒl] ��(Width�v���p�e�B)
'*****************************************************************************
Private Function GetWidth(ByRef objRange As Range) As Double
    Dim lngCount   As Long
    Dim lngHalf    As Long
    Dim MaxWidth   As Double '���̍ő�l

    MaxWidth = 32767 * DPIRatio
    If objRange.Width < MaxWidth Then
        GetWidth = objRange.Width
    Else
        With objRange
            '�O���{�㔼�̕������v
            lngCount = .Columns.Count
            lngHalf = lngCount / 2
            GetWidth = GetWidth(Range(.Columns(1), .Columns(lngHalf))) + _
                       GetWidth(Range(.Columns(lngHalf + 1), .Columns(lngCount)))
        End With
    End If
End Function


