Attribute VB_Name = "PixellColor"
Option Explicit

Public Type TRGBQuad
    Blue    As Byte
    Green   As Byte
    Red     As Byte
    Alpha   As Byte
End Type

Public Type TRGB
    Red     As Byte
    Green   As Byte
    Blue    As Byte
    None    As Byte
End Type

'�^�L���X�g�p
Private Type TLong
    Long   As Long
End Type

'*****************************************************************************
'[�T�v] Cell�̐F���擾����
'[����] �Ώۂ̃Z��
'[�ߒl] TRGBQuad
'*****************************************************************************
Public Function CellToColor(ByRef objCell As Range) As TRGBQuad
    Dim strNumeric As String
    Dim Alpha As Byte
    With objCell.Interior
        If .ColorIndex = xlColorIndexAutomatic Or .ColorIndex = xlColorIndexNone Then
            '����
            CellToColor = OleColorToARGB(&HFFFFFF, 0)
        Else
            Alpha = &HFF '�s����
            '������
            If .Pattern = xlGray8 Then
                strNumeric = Replace(objCell.Value, "$", "&H", 1, 1)
                If IsNumeric(strNumeric) Then
                    If 0 <= CInt(strNumeric) And CInt(strNumeric) <= 256 Then
                        '�Z���ɓ��͂��ꂽ���l���A���t�@�l
                        Alpha = CInt(strNumeric)
                    End If
                End If
            End If
            CellToColor = OleColorToARGB(.Color, Alpha)
        End If
    End With
End Function

'*****************************************************************************
'[�T�v] Cell�̐F��ݒ肷��
'[����] �F��ݒ肷��Z���C�ݒ肷��F
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub ColorToCell(ByRef objCell As Range, ByRef Color As TRGBQuad)
    With objCell.Interior
        Select Case Color.Alpha
        Case 0   '����
            .Pattern = xlGray8
            .PatternColorIndex = xlAutomatic
            .TintAndShade = 0
            .PatternTintAndShade = 0
            .ColorIndex = xlAutomatic
            objCell.Value = ""
        Case 255 '�s����
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .TintAndShade = 0
            .PatternTintAndShade = 0
            .Color = ARGBToOleColor(Color)
            objCell.Value = ""
        Case Else '������
            .Pattern = xlGray8
            .PatternColor = &HFFFFFF
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .PatternTintAndShade = 0
            .Color = ARGBToOleColor(Color)
            With objCell
                .Value = Color.Alpha
                .Font.Color = ARGBToOleColor(Color) '������w�i�F�Ɠ����ɂ���
            End With
        End Select
    End With
End Sub

'*****************************************************************************
'[�T�v] TRGBQuad�^��OLE_COLOR�ɕϊ�����
'[����] TRGBQuad
'[�ߒl] OLE_COLOR
'*****************************************************************************
Public Function ARGBToOleColor(ByRef ARGB As TRGBQuad) As OLE_COLOR
    Dim Color As TLong
    Dim RGB As TRGB
    With RGB
        .Red = ARGB.Red
        .Green = ARGB.Green
        .Blue = ARGB.Blue
    End With
    LSet Color = RGB
    ARGBToOleColor = Color.Long
End Function

'*****************************************************************************
'[�T�v] OLE_COLOR��TRGBQuad�^�ɕϊ�����
'[����] OLE_COLOR�C�A���t�@�l(�ȗ����͓��߂Ȃ�)
'[�ߒl] TRGBQuad
'*****************************************************************************
Public Function OleColorToARGB(ByVal lngColor As OLE_COLOR, Optional Alpha As Byte = 255) As TRGBQuad
    Dim RGB As TRGB
    Dim Color As TLong
    Color.Long = lngColor
    LSet RGB = Color
    With RGB
        OleColorToARGB = RGBToARGB(.Red, .Green, .Blue, Alpha)
    End With
End Function

'*****************************************************************************
'[�T�v] RGB & �A���t�@�l��TRGBQuad�^�ɕϊ�����
'[����] RGB�C�A���t�@�l(�ȗ����͓��߂Ȃ�)
'[�ߒl] TRGBQuad
'*****************************************************************************
Private Function RGBToARGB(ByVal Red As Byte, ByVal Green As Byte, ByVal Blue As Byte, Optional Alpha As Byte = 255) As TRGBQuad
    With RGBToARGB
        .Red = Red
        .Green = Green
        .Blue = Blue
        .Alpha = Alpha
    End With
End Function

'*****************************************************************************
'[�T�v] RGBQuad����v���邩�ǂ�������
'[����] ��r����F
'[�ߒl] True:��v
'*****************************************************************************
Public Function SameColor(ByRef RGBQuad1 As TRGBQuad, ByRef RGBQuad2 As TRGBQuad) As Boolean
    SameColor = (CastARGB(RGBQuad1) = CastARGB(RGBQuad2))
End Function

'*****************************************************************************
'[�T�v] RGB����уA���t�@�l�𑝌�������
'[����] SrcColor:�ύX�O�̐F�AlngUp:�����l(�}�C�i�X�͌����l)�ARGB���̂�����𑝌��̑ΏۂƂ��邩�ǂ���
'[�ߒl] �ύX��̐F
'*****************************************************************************
Public Function AdjustColor(ByRef SrcColor As TRGBQuad, ByVal lngUp As Long, ByVal blnRed As Boolean, ByVal blnGreen As Boolean, ByVal blnBlue As Boolean, ByVal blnAlpha As Boolean) As TRGBQuad
    AdjustColor = SrcColor
    With AdjustColor
        If lngUp > 0 Then
            If blnRed Then
                .Red = WorksheetFunction.Min(255, SrcColor.Red + lngUp)
            End If
            If blnGreen Then
                .Green = WorksheetFunction.Min(255, SrcColor.Green + lngUp)
            End If
            If blnBlue Then
                .Blue = WorksheetFunction.Min(255, SrcColor.Blue + lngUp)
            End If
            If blnAlpha Then
                .Alpha = WorksheetFunction.Min(255, SrcColor.Alpha + lngUp)
            End If
        Else
            If blnRed Then
                .Red = WorksheetFunction.MAX(0, SrcColor.Red + lngUp)
            End If
            If blnGreen Then
                .Green = WorksheetFunction.MAX(0, SrcColor.Green + lngUp)
            End If
            If blnBlue Then
                .Blue = WorksheetFunction.MAX(0, SrcColor.Blue + lngUp)
            End If
            If blnAlpha Then
                .Alpha = WorksheetFunction.MAX(0, SrcColor.Alpha + lngUp)
            End If
        End If
    End With
End Function

'*****************************************************************************
'[�T�v] TRGBQuad�^��Long�^�ɃL���X�g����(GDI+�̊֐��̈����ɓn������)
'[����] TRGBQuad
'[�ߒl] Long�^
'*****************************************************************************
Public Function CastARGB(ByRef ARGB As TRGBQuad) As Long
    Dim Color As TLong
'    Dim RGB As TRGB
'    With RGB
'        .Red = ARGB.Red
'        .Green = ARGB.Green
'        .Blue = ARGB.Blue
'    End With
    LSet Color = ARGB
    CastARGB = Color.Long
End Function

