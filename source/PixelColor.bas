Attribute VB_Name = "PixelColor"
Option Explicit
Option Private Module

Public Type TRGB
    Red     As Byte
    Green   As Byte
    Blue    As Byte
    None    As Byte
End Type

'型キャスト用
Public Type TLong
    Long   As Long
End Type

'*****************************************************************************
'[概要] Cellの色を取得する
'[引数] 対象のセル
'[戻値] TRGBQuad
'*****************************************************************************
Public Function CellToColor(ByRef objCell As Range) As TRGBQuad
    Dim Alpha As Byte
    Dim vValue As Variant
    With objCell.Interior
        Select Case .ColorIndex
        Case xlNone, xlAutomatic
            '透明
            CellToColor = OleColorToARGB(&HFFFFFF, 0)
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
            CellToColor = OleColorToARGB(.Color, Alpha)
        End Select
    End With
End Function

'*****************************************************************************
'[概要] Cellの色を設定する
'[引数] 色を設定するセル，設定する色，初期化が必要かどうか
'[戻値] なし
'*****************************************************************************
Public Sub ColorToCell(ByRef objCell As Range, ByRef Color As TRGBQuad, blnClear As Boolean)
    If objCell Is Nothing Then Exit Sub
    If blnClear Then
        Call ClearRange(objCell)
    End If
    With objCell.Interior
        Select Case Color.Alpha
        Case 0   '透明
            .Pattern = xlGray8
        Case 255 '不透明
            .Color = ARGBToOleColor(Color)
        Case Else '半透明
            .Color = ARGBToOleColor(Color)
            .Pattern = xlGray8
            .PatternColor = &HFFFFFF '白
            objCell.Value = Color.Alpha
            objCell.Font.Color = .Color  '文字を背景色と同じにする
        End Select
    End With
End Sub

'*****************************************************************************
'[概要] TRGBQuad型をOLE_COLORに変換する
'[引数] TRGBQuad
'[戻値] OLE_COLOR
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
'[概要] OLE_COLORをTRGBQuad型に変換する
'[引数] OLE_COLOR，アルファ値(省略時は透過なし)
'[戻値] TRGBQuad
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
'[概要] RGB & アルファ値をTRGBQuad型に変換する
'[引数] RGB，アルファ値(省略時は透過なし)
'[戻値] TRGBQuad
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
'[概要] RGBQuadが一致するかどうか判定
'[引数] 比較する色
'[戻値] True:一致
'*****************************************************************************
Public Function SameColor(ByRef RGBQuad1 As TRGBQuad, ByRef RGBQuad2 As TRGBQuad) As Boolean
    SameColor = (CastARGB(RGBQuad1) = CastARGB(RGBQuad2))
End Function

'*****************************************************************************
'[概要] RGBおよびアルファ値を増減させる
'[引数] SrcColor:変更前の色、RGBαのそれぞれの増減値
'[戻値] 変更後の色
'*****************************************************************************
Public Function AdjustColor(ByRef SrcColor As TRGBQuad, ByVal Red As Long, ByVal Green As Long, ByVal Blue As Long, ByVal Alpha As Long) As TRGBQuad
    AdjustColor = SrcColor
    With AdjustColor
        If Red + Green + Blue + Alpha > 0 Then
            '増加の時
            .Red = WorksheetFunction.min(255, SrcColor.Red + Red)
            .Green = WorksheetFunction.min(255, SrcColor.Green + Green)
            .Blue = WorksheetFunction.min(255, SrcColor.Blue + Blue)
            .Alpha = WorksheetFunction.min(255, SrcColor.Alpha + Alpha)
        Else
            '減少の時
            .Red = WorksheetFunction.max(0, SrcColor.Red + Red)
            .Green = WorksheetFunction.max(0, SrcColor.Green + Green)
            .Blue = WorksheetFunction.max(0, SrcColor.Blue + Blue)
            .Alpha = WorksheetFunction.max(0, SrcColor.Alpha + Alpha)
        End If
    End With
End Function

'*****************************************************************************
'[概要] 色のHSLを増減させる
'[引数] SrcColor:変更前の色、HSLのそれぞれの増減値
'[戻値] 変更後の色
'*****************************************************************************
Public Function UpDownHSL(ByRef SrcColor As TRGBQuad, ByVal Hue As Long, ByVal Saturation As Long, ByVal Lightness As Long, Optional LeVel As Long = 1) As TRGBQuad
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
    With UpDownHSL
        .Red = Round(R)
        .Green = Round(G)
        .Blue = Round(B)
        .Alpha = SrcColor.Alpha
    End With
    If Not SameColor(SrcColor, UpDownHSL) Then Exit Function
    
    '変更前後で値が一致した場合は、不一致するまで再帰呼び出し
    If LeVel = 3 Then
        Exit Function
    End If
    LeVel = LeVel + 1
    UpDownHSL = UpDownHSL(UpDownHSL, Hue * LeVel, Saturation * LeVel, Lightness * LeVel, LeVel)
    
    '変更前後で値が一致した場合は、乖離の一番大きい色を補正する
'    Dim blnMaxValue(1 To 3) As Boolean
'    If Hue + Saturation + Lightness > 0 Then
'        blnMaxValue(1) = (SrcColor.Red = 255)
'        blnMaxValue(2) = (SrcColor.Green = 255)
'        blnMaxValue(3) = (SrcColor.Blue = 255)
'    Else
'        '減少の時
'        blnMaxValue(1) = (SrcColor.Red = 0)
'        blnMaxValue(2) = (SrcColor.Green = 0)
'        blnMaxValue(3) = (SrcColor.Blue = 0)
'    End If
'
'    Dim i As Long
'    Dim Diff(1 To 3) As Double
'    If Not blnMaxValue(1) Then
'        Diff(1) = R - SrcColor.Red
'    End If
'    If Not blnMaxValue(2) Then
'        Diff(2) = G - SrcColor.Green
'    End If
'    If Not blnMaxValue(3) Then
'        Diff(3) = B - SrcColor.Blue
'    End If
'
'    Dim sign(1 To 3) As Long '符号
'    For i = 1 To 3
'        If Diff(i) < 0 Then
'            sign(i) = -1
'            Diff(i) = Abs(Diff(i))
'        Else
'            sign(i) = 1
'        End If
'    Next
'
'    Dim MaxDiff As Double
'    MaxDiff = WorksheetFunction.max(Diff(1), Diff(2), Diff(3))
'    If MaxDiff = 0 Then
'        '真っ白、または真っ黒の時
'        Exit Function
'    End If
'    If MaxDiff = Diff(1) Then
'       UpDownHSL.Red = UpDownHSL.Red + sign(1)
'    ElseIf MaxDiff = Diff(2) Then
'       UpDownHSL.Green = UpDownHSL.Green + sign(2)
'    Else
'       UpDownHSL.Blue = UpDownHSL.Blue + sign(3)
'    End If
End Function

'*****************************************************************************
'[概要] RGBをHSLに変換する
'[引数] SrcColor:変更前の色, 計算結果：H:0～360,S:0～255,L:0～255
'[戻値] 変換後のHSL(ただし引数)
'*****************************************************************************
Private Sub RGBToHSL(ByRef SrcColor As TRGBQuad, ByRef H As Double, ByRef S As Double, ByRef L As Double)
    Dim R As Long '0～255
    Dim G As Long '0～255
    Dim B As Long '0～255
    With SrcColor
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
'[概要] TRGBQuad型をLong型にキャストする(GDI+の関数の引数に渡すため)
'[引数] TRGBQuad
'[戻値] Long型
'*****************************************************************************
Public Function CastARGB(ByRef ARGB As TRGBQuad) As Long
    Dim Color As TLong
    LSet Color = ARGB
    CastARGB = Color.Long
End Function

