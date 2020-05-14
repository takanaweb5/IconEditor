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
        Select Case .ColorIndex
        Case xlNone, xlAutomatic
            '����
            CellToColor = OleColorToARGB(&HFFFFFF, 0)
        Case Else
            Alpha = &HFF '�s����
            '���������ǂ���
            If .Pattern = xlGray8 Then
                strNumeric = Replace(objCell.Value, "$", "&H", 1, 1)
                If IsNumeric(strNumeric) Then
                    If 0 <= CLng(strNumeric) And CLng(strNumeric) <= 255 Then
                        '�Z���ɓ��͂��ꂽ���l���A���t�@�l
                        Alpha = CByte(strNumeric)
                    End If
                End If
            End If
            CellToColor = OleColorToARGB(.Color, Alpha)
        End Select
    End With
End Function

'*****************************************************************************
'[�T�v] Cell�̐F��ݒ肷��
'[����] �F��ݒ肷��Z���C�ݒ肷��F�C���������K�v���ǂ���
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub ColorToCell(ByRef objCell As Range, ByRef Color As TRGBQuad, blnClear As Boolean)
    If objCell Is Nothing Then Exit Sub
    If blnClear Then
        Call ClearRange(objCell)
    End If
    With objCell.Interior
        Select Case Color.Alpha
        Case 0   '����
            .Pattern = xlGray8
        Case 255 '�s����
            .Color = ARGBToOleColor(Color)
        Case Else '������
            .Color = ARGBToOleColor(Color)
            .Pattern = xlGray8
            .PatternColor = &HFFFFFF '��
            objCell.Value = Color.Alpha
            objCell.Font.Color = ARGBToOleColor(Color) '������w�i�F�Ɠ����ɂ���
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
'[����] SrcColor:�ύX�O�̐F�ARGB���̂��ꂼ��̑����l
'[�ߒl] �ύX��̐F
'*****************************************************************************
Public Function AdjustColor(ByRef SrcColor As TRGBQuad, ByVal Red As Long, ByVal Green As Long, ByVal Blue As Long, ByVal Alpha As Long) As TRGBQuad
    AdjustColor = SrcColor
    With AdjustColor
        If Red + Green + Blue + Alpha > 0 Then
            '�����̎�
            .Red = WorksheetFunction.min(255, SrcColor.Red + Red)
            .Green = WorksheetFunction.min(255, SrcColor.Green + Green)
            .Blue = WorksheetFunction.min(255, SrcColor.Blue + Blue)
            .Alpha = WorksheetFunction.min(255, SrcColor.Alpha + Alpha)
        Else
            '�����̎�
            .Red = WorksheetFunction.max(0, SrcColor.Red + Red)
            .Green = WorksheetFunction.max(0, SrcColor.Green + Green)
            .Blue = WorksheetFunction.max(0, SrcColor.Blue + Blue)
            .Alpha = WorksheetFunction.max(0, SrcColor.Alpha + Alpha)
        End If
    End With
End Function

'*****************************************************************************
'[�T�v] �F��HSL�𑝌�������
'[����] SrcColor:�ύX�O�̐F�AHSL�̂��ꂼ��̑����l
'[�ߒl] �ύX��̐F
'*****************************************************************************
Public Function UpDownHSL(ByRef SrcColor As TRGBQuad, ByVal Hue As Long, ByVal Saturation As Long, ByVal Lightness As Long) As TRGBQuad
    Dim H As Double '0�`360
    Dim S As Double '0�`255
    Dim L As Double '0�`255
        
    Call RGBToHSL(SrcColor, H, S, L)
    If Hue + Saturation + Lightness > 0 Then
        '�����̎�
        H = H + Hue
        S = WorksheetFunction.min(255, S + Saturation)
        L = WorksheetFunction.min(255, L + Lightness)
    Else
        '�����̎�
        H = H + Hue
        S = WorksheetFunction.max(0, S + Saturation)
        L = WorksheetFunction.max(0, L + Lightness)
    End If
    UpDownHSL = HSLToRGB(H, S, L)
    UpDownHSL.Alpha = SrcColor.Alpha
End Function

'*****************************************************************************
'[�T�v] RGB��HSL�ɕϊ�����
'[����] SrcColor:�ύX�O�̐F, �v�Z���ʁFH:0�`360,S:0�`255,L:0�`255
'[�ߒl] �ϊ����HSL(����������)
'*****************************************************************************
Private Sub RGBToHSL(ByRef SrcColor As TRGBQuad, ByRef H As Double, ByRef S As Double, ByRef L As Double)
    Dim R As Long '0�`255
    Dim G As Long '0�`255
    Dim B As Long '0�`255
    With SrcColor
        R = .Red
        G = .Green
        B = .Blue
    End With
    
    Dim max As Long
    Dim min As Long
    max = WorksheetFunction.max(R, G, B)
    min = WorksheetFunction.min(R, G, B)
    
    'L(���x)
    L = (max + min) / 2
    
    'H(�F��)
    If max <> min Then
        If max = R Then
            H = 60 * (G - B) / (max - min)
        End If
        If max = G Then
            H = 60 * (B - R) / (max - min) + 120
        End If
        If max = B Then
            H = 60 * (R - G) / (max - min) + 240
        End If
        If H < 0 Then
            H = H + 360
        End If
    End If
     
    'S(�ʓx)
    If max <> min Then
        If L <= 127 Then
          S = (max - min) / (max + min)
        Else
          S = (max - min) / (510 - max - min)
        End If
        S = S * 255
    End If
End Sub

'*****************************************************************************
'[�T�v] HSL��RGB�ɕϊ�����
'[����] H:0�`360,S:0�`255,L:0�`255
'[�ߒl] RGB
'*****************************************************************************
Private Function HSLToRGB(ByVal H As Double, ByVal S As Double, ByVal L As Double) As TRGBQuad
    Dim max As Double
    Dim min As Double
    If L <= 127 Then
      max = L + L * (S / 255)
      min = L - L * (S / 255)
    Else
      max = (L + (255 - L) * (S / 255))
      min = (L - (255 - L) * (S / 255))
    End If

    Dim R As Double '0�`255
    Dim G As Double '0�`255
    Dim B As Double '0�`255
    If H < 0 Then
        H = H + 360
    End If
    If H >= 360 Then
        H = H - 360
    End If
    If H < 60 Then
        R = max
        G = min + (max - min) * (H / 60)
        B = min
    ElseIf 60 <= H And H < 120 Then
        R = min + (max - min) * ((120 - H) / 60)
        G = max
        B = min
    ElseIf 120 <= H And H < 180 Then
        R = min
        G = max
        B = min + (max - min) * ((H - 120) / 60)
    ElseIf 180 <= H And H < 240 Then
        R = min
        G = min + (max - min) * ((240 - H) / 60)
        B = max
    ElseIf 240 <= H And H < 300 Then
        R = min + (max - min) * ((H - 240) / 60)
        G = min
        B = max
    ElseIf 300 <= H And H < 360 Then
        R = max
        G = min
        B = min + (max - min) * ((360 - H) / 60)
    End If
    With HSLToRGB
        .Red = Round(R)
        .Green = Round(G)
        .Blue = Round(B)
    End With
End Function

'*****************************************************************************
'[�T�v] TRGBQuad�^��Long�^�ɃL���X�g����(GDI+�̊֐��̈����ɓn������)
'[����] TRGBQuad
'[�ߒl] Long�^
'*****************************************************************************
Public Function CastARGB(ByRef ARGB As TRGBQuad) As Long
    Dim Color As TLong
    LSet Color = ARGB
    CastARGB = Color.Long
End Function

