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

'型キャスト用
Private Type TLong
    Long   As Long
End Type

'*****************************************************************************
'[概要] Cellの色を取得する
'[引数] 対象のセル
'[戻値] TRGBQuad
'*****************************************************************************
Public Function CellToColor(ByRef objCell As Range) As TRGBQuad
    Dim strNumeric As String
    Dim Alpha As Byte
    With objCell.Interior
        If .ColorIndex = xlColorIndexAutomatic Or .ColorIndex = xlColorIndexNone Then
            '透明
            CellToColor = OleColorToARGB(&HFFFFFF, 0)
        Else
            Alpha = &HFF '不透明
            '半透明
            If .Pattern = xlGray8 Then
                strNumeric = Replace(objCell.Value, "$", "&H", 1, 1)
                If IsNumeric(strNumeric) Then
                    If 0 <= CInt(strNumeric) And CInt(strNumeric) <= 256 Then
                        'セルに入力された数値がアルファ値
                        Alpha = CInt(strNumeric)
                    End If
                End If
            End If
            CellToColor = OleColorToARGB(.Color, Alpha)
        End If
    End With
End Function

'*****************************************************************************
'[概要] Cellの色を設定する
'[引数] 色を設定するセル，設定する色
'[戻値] なし
'*****************************************************************************
Public Sub ColorToCell(ByRef objCell As Range, ByRef Color As TRGBQuad)
    With objCell.Interior
        Select Case Color.Alpha
        Case 0   '透明
            .Pattern = xlGray8
            .PatternColorIndex = xlAutomatic
            .TintAndShade = 0
            .PatternTintAndShade = 0
            .ColorIndex = xlAutomatic
            objCell.Value = ""
        Case 255 '不透明
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .TintAndShade = 0
            .PatternTintAndShade = 0
            .Color = ARGBToOleColor(Color)
            objCell.Value = ""
        Case Else '半透明
            .Pattern = xlGray8
            .PatternColor = &HFFFFFF
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .PatternTintAndShade = 0
            .Color = ARGBToOleColor(Color)
            With objCell
                .Value = Color.Alpha
                .Font.Color = ARGBToOleColor(Color) '文字を背景色と同じにする
            End With
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
'[引数] SrcColor:変更前の色、lngUp:増加値(マイナスは減少値)、RGBαのいずれを増減の対象とするかどうか
'[戻値] 変更後の色
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
'[概要] TRGBQuad型をLong型にキャストする(GDI+の関数の引数に渡すため)
'[引数] TRGBQuad
'[戻値] Long型
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

