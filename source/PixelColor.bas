Attribute VB_Name = "PixelColor"
Option Explicit
Option Private Module

Public Type TRGB
    Red     As Byte
    Green   As Byte
    Blue    As Byte
    None    As Byte
End Type

Public Type TRGBQuad
    Blue    As Byte
    Green   As Byte
    Red     As Byte
    Alpha   As Byte
End Type

Public Const CTRANSPARENT = &HFFFFFF

'�^�L���X�g�p
Public Type TLong
    Long   As Long
End Type

'*****************************************************************************
'[�T�v] Cell�̐F���擾����
'[����] �Ώۂ̃Z��
'[�ߒl] RGBQuad
'*****************************************************************************
Public Function CellToRGBQuad(ByRef objCell As Range) As Long
    Dim Alpha As Byte
    Dim vValue As Variant
    With objCell.Interior
        Select Case .ColorIndex
        Case xlNone, xlAutomatic
            '����
            CellToRGBQuad = CTRANSPARENT
        Case Else
            Alpha = &HFF '�s����
            '���������ǂ���
            If .Pattern = xlGray8 Then
                vValue = objCell.Value
                If IsNumeric(vValue) And vValue <> "" Then
                    If 0 <= CLng(vValue) And CLng(vValue) <= 255 Then
                        '�Z���ɓ��͂��ꂽ���l���A���t�@�l
                        Alpha = CByte(vValue)
                    End If
                End If
            End If
            CellToRGBQuad = OleColorToRGBQuad(.Color, Alpha)
        End Select
    End With
End Function

'*****************************************************************************
'[�T�v] Cell�̐F��ݒ肷��
'[����] �F��ݒ肷��Z���C�ݒ肷��F(RGBQuad)�C���������K�v���ǂ���
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub RGBQuadToCell(ByRef objCell As Range, ByVal RGBQuad As Long, ByVal blnClear As Boolean)
    If objCell Is Nothing Then Exit Sub
    If blnClear Then
        Call ClearRange(objCell)
    End If
    With objCell.Interior
        Select Case RGBQuadToAlpha(RGBQuad)
        Case 0   '����
            .Pattern = xlGray8
        Case 255 '�s����
            .Color = RGBQuadToOleColor(RGBQuad)
        Case Else '������
            .Color = RGBQuadToOleColor(RGBQuad)
            .Pattern = xlGray8
            .PatternColor = &HFFFFFF '��
            objCell.Value = RGBQuadToAlpha(RGBQuad)
            objCell.Font.Color = .Color  '������w�i�F�Ɠ����ɂ���
        End Select
    End With
End Sub

'*****************************************************************************
'[�T�v] RGBQuad��OLE_COLOR�ɕϊ�����
'[����] RGBQuad
'[�ߒl] OLE_COLOR
'*****************************************************************************
Public Function RGBQuadToOleColor(ByVal lngRGBQuad As Long) As OLE_COLOR
    Dim Color   As TLong
    Dim RGBQuad As TRGBQuad
    Dim RGB     As TRGB
    Color.Long = lngRGBQuad
    LSet RGBQuad = Color
    With RGB
        .Red = RGBQuad.Red
        .Green = RGBQuad.Green
        .Blue = RGBQuad.Blue
    End With
    LSet Color = RGB
    RGBQuadToOleColor = Color.Long
End Function

'*****************************************************************************
'[�T�v] OLE_COLOR��RGBQuad�ɕϊ�����
'[����] OLE_COLOR�C�A���t�@�l(�ȗ����͓��߂Ȃ�)
'[�ߒl] RGBQuad
'*****************************************************************************
Public Function OleColorToRGBQuad(ByVal lngColor As OLE_COLOR, Optional Alpha As Byte = 255) As Long
    Dim Color   As TLong
    Dim RGBQuad As TRGBQuad
    Dim RGB     As TRGB
    Color.Long = lngColor
    LSet RGB = Color
    With RGBQuad
        .Red = RGB.Red
        .Green = RGB.Green
        .Blue = RGB.Blue
        .Alpha = Alpha
    End With
    LSet Color = RGBQuad
    OleColorToRGBQuad = Color.Long
End Function

'*****************************************************************************
'[�T�v] RGB & �A���t�@�l��RGBQuad�ɕϊ�����
'[����] RGB�C�A���t�@�l(�ȗ����͓��߂Ȃ�)
'[�ߒl] TRGBQuad
'*****************************************************************************
Private Function RGBToRGBQuad(ByVal Red As Byte, ByVal Green As Byte, ByVal Blue As Byte, Optional Alpha As Byte = 255) As Long
    Dim Color   As TLong
    Dim RGBQuad As TRGBQuad
    With RGBQuad
        .Red = Red
        .Green = Green
        .Blue = Blue
        .Alpha = Alpha
    End With
    LSet Color = RGBQuad
    RGBToRGBQuad = Color.Long
End Function

'*****************************************************************************
'[�T�v] RGB����уA���t�@�l�𑝌�������
'[����] SrcColor:�ύX�O�̐F�ARGB���̂��ꂼ��̑����l
'[�ߒl] �ύX��̐F
'*****************************************************************************
Public Function AdjustColor(ByVal SrcColor As Long, ByVal Red As Long, ByVal Green As Long, ByVal Blue As Long, ByVal Alpha As Long) As Long
    Dim RGBQuad As TRGBQuad
    Dim Color   As TLong
    Color.Long = SrcColor
    LSet RGBQuad = Color
    With RGBQuad
        If Red + Green + Blue + Alpha > 0 Then
            '�����̎�
            .Red = WorksheetFunction.min(255, .Red + Red)
            .Green = WorksheetFunction.min(255, .Green + Green)
            .Blue = WorksheetFunction.min(255, .Blue + Blue)
            .Alpha = WorksheetFunction.min(255, .Alpha + Alpha)
        Else
            '�����̎�
            .Red = WorksheetFunction.max(0, .Red + Red)
            .Green = WorksheetFunction.max(0, .Green + Green)
            .Blue = WorksheetFunction.max(0, .Blue + Blue)
            .Alpha = WorksheetFunction.max(0, .Alpha + Alpha)
        End If
    End With
    LSet Color = RGBQuad
    AdjustColor = Color.Long
End Function

'*****************************************************************************
'[�T�v] �F��HSL�𑝌�������
'[����] SrcColor:�ύX�O�̐F�AHSL�̂��ꂼ��̑����l
'[�ߒl] �ύX��̐F
'*****************************************************************************
Public Function UpDownHSL(ByVal SrcColor As Long, ByVal Hue As Long, ByVal Saturation As Long, ByVal Lightness As Long, Optional LeVel As Long = 1) As Long
    Dim H As Double '0�`360
    Dim S As Double '0�`255
    Dim L As Double '0�`255
        
    Call RGBToHSL(SrcColor, H, S, L)
    
'    Debug.Print LeVel & "  " & "R:" & SrcColor.Red & " " & "G:" & SrcColor.Green & " " & "B:" & SrcColor.Blue, _
'                "H:" & H & " " & "S:" & S & " " & "L:" & L

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
    Dim R As Double
    Dim G As Double
    Dim B As Double
    Call HSLToRGB(H, S, L, R, G, B)
    UpDownHSL = RGBToRGBQuad(Round(R), Round(G), Round(B), RGBQuadToAlpha(SrcColor))
    If SrcColor <> UpDownHSL Then Exit Function
    
    '�ύX�O��Œl����v�����ꍇ�́A�s��v����܂ōċA�Ăяo��
    If LeVel = 3 Then
        Exit Function
    End If
    LeVel = LeVel + 1
    UpDownHSL = UpDownHSL(UpDownHSL, Hue * LeVel, Saturation * LeVel, Lightness * LeVel, LeVel)
End Function

'*****************************************************************************
'[�T�v] RGB��HSL�ɕϊ�����
'[����] SrcColor:�ύX�O�̐F, �v�Z���ʁFH:0�`360,S:0�`255,L:0�`255
'[�ߒl] �ϊ����HSL(����������)
'*****************************************************************************
Private Sub RGBToHSL(ByVal SrcColor As Long, ByRef H As Double, ByRef S As Double, ByRef L As Double)
    Dim R As Long '0�`255
    Dim G As Long '0�`255
    Dim B As Long '0�`255
    
    Dim RGBQuad As TRGBQuad
    Dim Color   As TLong
    Color.Long = SrcColor
    LSet RGBQuad = Color
    With RGBQuad
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
'[�ߒl] RGB(��4�����`��6����)
'*****************************************************************************
Private Sub HSLToRGB(ByVal H As Double, ByVal S As Double, ByVal L As Double, ByRef R As Double, ByRef G As Double, ByRef B As Double)
    Dim max As Double
    Dim min As Double
    If L <= 127 Then
      max = L + L * (S / 255)
      min = L - L * (S / 255)
    Else
      max = (L + (255 - L) * (S / 255))
      min = (L - (255 - L) * (S / 255))
    End If

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
End Sub

'*****************************************************************************
'[�T�v] TRGBQuad�^�̃��l���擾����
'[����] TRGBQuad
'[�ߒl] ���l
'*****************************************************************************
Public Function RGBQuadToAlpha(ByVal lngRGBQuad As Long) As Byte
    Dim Color As TLong
    Dim RGBQuad As TRGBQuad
    Color.Long = lngRGBQuad
    LSet RGBQuad = Color
    RGBQuadToAlpha = RGBQuad.Alpha
End Function

