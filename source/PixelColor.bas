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

'型キャスト用
Public Type TLong
    Long   As Long
End Type

'*****************************************************************************
'[概要] Cellの色を取得する
'[引数] 対象のセル
'[戻値] RGBQuad
'*****************************************************************************
Public Function CellToRGBQuad(ByRef objCell As Range) As Long
    Dim Alpha As Byte
    Dim vValue As Variant
    With objCell.Interior
        Select Case .ColorIndex
        Case xlNone, xlAutomatic
            '透明
            CellToRGBQuad = CTRANSPARENT
        Case Else
            Alpha = &HFF '不透明
            '半透明かどうか
            If .Pattern = xlGray8 Then
                vValue = objCell.Value
                If IsNumeric(vValue) And vValue <> "" Then
                    If 0 <= CLng(vValue) And CLng(vValue) <= 255 Then
                        'セルに入力された数値がアルファ値
                        Alpha = CByte(vValue)
                    End If
                End If
            End If
            CellToRGBQuad = OleColorToRGBQuad(.Color, Alpha)
        End Select
    End With
End Function

'*****************************************************************************
'[概要] Cellの色を設定する
'[引数] 色を設定するセル，設定する色(RGBQuad)，初期化が必要かどうか
'[戻値] なし
'*****************************************************************************
Public Sub RGBQuadToCell(ByRef objCell As Range, ByVal RGBQuad As Long, ByVal blnClear As Boolean)
    If objCell Is Nothing Then Exit Sub
    If blnClear Then
        Call ClearRange(objCell)
    End If
    With objCell.Interior
        Select Case RGBQuadToAlpha(RGBQuad)
        Case 0   '透明
            .Pattern = xlGray8
        Case 255 '不透明
            .Color = RGBQuadToOleColor(RGBQuad)
        Case Else '半透明
            .Color = RGBQuadToOleColor(RGBQuad)
            .Pattern = xlGray8
            .PatternColor = &HFFFFFF '白
            objCell.Value = RGBQuadToAlpha(RGBQuad)
            objCell.Font.Color = .Color  '文字を背景色と同じにする
        End Select
    End With
End Sub

'*****************************************************************************
'[概要] RGBQuadをOLE_COLORに変換する
'[引数] RGBQuad
'[戻値] OLE_COLOR
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
'[概要] OLE_COLORをRGBQuadに変換する
'[引数] OLE_COLOR，アルファ値(省略時は透過なし)
'[戻値] RGBQuad
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
'[概要] RGB & アルファ値をRGBQuadに変換する
'[引数] RGB，アルファ値(省略時は透過なし)
'[戻値] TRGBQuad
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
'[概要] RGBおよびアルファ値を増減させる
'[引数] SrcColor:変更前の色、RGBαのそれぞれの増減値
'[戻値] 変更後の色
'*****************************************************************************
Public Function AdjustColor(ByVal SrcColor As Long, ByVal Red As Long, ByVal Green As Long, ByVal Blue As Long, ByVal Alpha As Long) As Long
    Dim RGBQuad As TRGBQuad
    Dim Color   As TLong
    Color.Long = SrcColor
    LSet RGBQuad = Color
    With RGBQuad
        If Red + Green + Blue + Alpha > 0 Then
            '増加の時
            .Red = WorksheetFunction.min(255, .Red + Red)
            .Green = WorksheetFunction.min(255, .Green + Green)
            .Blue = WorksheetFunction.min(255, .Blue + Blue)
            .Alpha = WorksheetFunction.min(255, .Alpha + Alpha)
        Else
            '減少の時
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
'[概要] 色のHSLを増減させる
'[引数] SrcColor:変更前の色、HSLのそれぞれの増減値
'[戻値] 変更後の色
'*****************************************************************************
Public Function UpDownHSL(ByVal SrcColor As Long, ByVal Hue As Long, ByVal Saturation As Long, ByVal Lightness As Long, Optional LeVel As Long = 1) As Long
    Dim H As Double '0～360
    Dim S As Double '0～255
    Dim L As Double '0～255
        
    Call RGBToHSL(SrcColor, H, S, L)
    
'    Debug.Print LeVel & "  " & "R:" & SrcColor.Red & " " & "G:" & SrcColor.Green & " " & "B:" & SrcColor.Blue, _
'                "H:" & H & " " & "S:" & S & " " & "L:" & L

    If Hue + Saturation + Lightness > 0 Then
        '増加の時
        H = H + Hue
        S = WorksheetFunction.min(255, S + Saturation)
        L = WorksheetFunction.min(255, L + Lightness)
    Else
        '減少の時
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
    
    '変更前後で値が一致した場合は、不一致するまで再帰呼び出し
    If LeVel = 3 Then
        Exit Function
    End If
    LeVel = LeVel + 1
    UpDownHSL = UpDownHSL(UpDownHSL, Hue * LeVel, Saturation * LeVel, Lightness * LeVel, LeVel)
End Function

'*****************************************************************************
'[概要] RGBをHSLに変換する
'[引数] SrcColor:変更前の色, 計算結果：H:0～360,S:0～255,L:0～255
'[戻値] 変換後のHSL(ただし引数)
'*****************************************************************************
Private Sub RGBToHSL(ByVal SrcColor As Long, ByRef H As Double, ByRef S As Double, ByRef L As Double)
    Dim R As Long '0～255
    Dim G As Long '0～255
    Dim B As Long '0～255
    
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
    
    'L(明度)
    L = (max + min) / 2
    
    'H(色相)
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
     
    'S(彩度)
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
'[概要] HSLをRGBに変換する
'[引数] H:0～360,S:0～255,L:0～255
'[戻値] RGB(第4引数～第6引数)
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
'[概要] TRGBQuad型のα値を取得する
'[引数] TRGBQuad
'[戻値] α値
'*****************************************************************************
Public Function RGBQuadToAlpha(ByVal lngRGBQuad As Long) As Byte
    Dim Color As TLong
    Dim RGBQuad As TRGBQuad
    Color.Long = lngRGBQuad
    LSet RGBQuad = Color
    RGBQuadToAlpha = RGBQuad.Alpha
End Function

