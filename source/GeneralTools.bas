Attribute VB_Name = "GeneralTools"
Option Explicit

Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
Private Declare PtrSafe Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As LongPtr
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As LongPtr
Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)

Private Const CF_BITMAP = 2

'�I���^�C�v
Public Enum ESelectionType
    E_Range
    E_Shape
    E_Non
    E_Other
End Enum

Public Const MAX_WIDTH = 64
Public Const MAX_HEIGHT = 64

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
    If objRange.Count = 1 Then
        Set ReSelectRange = objRange
        Exit Function
    End If
    
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



