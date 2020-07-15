Attribute VB_Name = "CellFunction"
Option Explicit
'Option Private Module

'*****************************************************************************
'[概要] セルに適用されるRGBAカラーを取得する
'[引数] 対象セル
'[戻値] RGBAの16進数(ただし透明は0)
'*****************************************************************************
Public Function Cell2RGBA(ByVal objCell As Range) As Variant
Attribute Cell2RGBA.VB_ProcData.VB_Invoke_Func = " \n14"
On Error GoTo ErrHandle
    Dim Alpha As String
    Dim vValue As Variant
    With objCell.Interior
        Select Case .ColorIndex
        Case xlNone, xlAutomatic
            '透明
            Cell2RGBA = 0
        Case Else
            Alpha = "FF" '不透明
            '半透明かどうか
            If .Pattern = xlGray8 Then
                vValue = objCell.Value
                If IsNumeric(vValue) Then
                    If 0 <= CLng(vValue) And CLng(vValue) <= 255 Then
                        'セルに入力された数値がアルファ値
                        Alpha = WorksheetFunction.Dec2Hex(vValue, 2)
                    End If
                End If
            End If
            Cell2RGBA = "${R}{G}{B}" & Alpha
            Dim Color As TLong
            Color.Long = OleColorToRGBQuad(.Color)
            Dim RGBQuad As TRGBQuad
            LSet RGBQuad = Color
            With RGBQuad
                Cell2RGBA = Replace(Cell2RGBA, "{R}", WorksheetFunction.Dec2Hex(.Red, 2))
                Cell2RGBA = Replace(Cell2RGBA, "{G}", WorksheetFunction.Dec2Hex(.Green, 2))
                Cell2RGBA = Replace(Cell2RGBA, "{B}", WorksheetFunction.Dec2Hex(.Blue, 2))
            End With
        End Select
    End With
Exit Function
ErrHandle:
    Cell2RGBA = ""
End Function

'*****************************************************************************
'[概要] セルに適用されるRGBカラーを取得する
'[引数] 対象セル(省略時はセル関数の設定されているセル)
'[戻値] RGBの16進数(ただし透明は0)
'*****************************************************************************
Public Function Cell2RGB(Optional objCell As Range = Nothing) As Variant
On Error GoTo ErrHandle
    Dim TargetCell As Range
    If objCell Is Nothing Then
        Set TargetCell = Application.ThisCell
    Else
        Set TargetCell = objCell
    End If
    
    With TargetCell.Interior
        Select Case .ColorIndex
        Case xlNone, xlAutomatic
            '透明
            Cell2RGB = 0
        Case Else
            Cell2RGB = "${R}{G}{B}"
            Dim Color As TLong
            Color.Long = OleColorToRGBQuad(.Color)
            Dim RGBQuad As TRGBQuad
            LSet RGBQuad = Color
            With RGBQuad
                Cell2RGB = Replace(Cell2RGB, "{R}", WorksheetFunction.Dec2Hex(.Red, 2))
                Cell2RGB = Replace(Cell2RGB, "{G}", WorksheetFunction.Dec2Hex(.Green, 2))
                Cell2RGB = Replace(Cell2RGB, "{B}", WorksheetFunction.Dec2Hex(.Blue, 2))
            End With
        End Select
    End With
Exit Function
ErrHandle:
    Cell2RGB = ""
End Function

'*****************************************************************************
'[概要] 16進数のRGBのRed部分を取得
'[引数] 16進数のRGBまたはRGBA
'[戻値] Redの値(0～255)
'*****************************************************************************
Public Function Hex2Red(ByVal strHex As String) As Byte
    If Left(strHex, 1) = "$" Then
        If Len(strHex) = 7 Or Len(strHex) = 9 Then
            If IsNumeric("&H" & Mid(strHex, 2, 2)) Then
                Hex2Red = "&H" & Mid(strHex, 2, 2)
                Exit Function
            End If
        End If
    End If
    '例外が発生し、#VALUE!となる
    Hex2Red = CVErr(xlErrValue)
End Function

'*****************************************************************************
'[概要] 16進数のRGBのGreen部分を取得
'[引数] 16進数のRGBまたはRGBA
'[戻値] Greenの値(0～255)
'*****************************************************************************
Public Function Hex2Green(ByVal strHex As String) As Byte
    If Left(strHex, 1) = "$" Then
        If Len(strHex) = 7 Or Len(strHex) = 9 Then
            If IsNumeric("&H" & Mid(strHex, 4, 2)) Then
                Hex2Green = "&H" & Mid(strHex, 4, 2)
                Exit Function
            End If
        End If
    End If
    '例外が発生し、#VALUE!となる
    Hex2Green = CVErr(xlErrValue)
End Function

'*****************************************************************************
'[概要] 16進数のRGBのBlue部分を取得
'[引数] 16進数のRGBまたはRGBA
'[戻値] Blueの値(0～255)
'*****************************************************************************
Public Function Hex2Blue(ByVal strHex As String) As Byte
    If Left(strHex, 1) = "$" Then
        If Len(strHex) = 7 Or Len(strHex) = 9 Then
            If IsNumeric("&H" & Mid(strHex, 6, 2)) Then
                Hex2Blue = "&H" & Mid(strHex, 6, 2)
                Exit Function
            End If
        End If
    End If
    '例外が発生し、#VALUE!となる
    Hex2Blue = CVErr(xlErrValue)
End Function

'*****************************************************************************
'[概要] 16進数のRGBを設定
'[引数] Red,Green,Blue
'[戻値] RGB2Hex(255,128,0)→$FF8000
'*****************************************************************************
Public Function RGB2Hex(ByVal Red As Byte, ByVal Green As Byte, ByVal Blue As Byte) As String
    RGB2Hex = "$" & WorksheetFunction.Dec2Hex(Red, 2) _
                  & WorksheetFunction.Dec2Hex(Green, 2) _
                  & WorksheetFunction.Dec2Hex(Blue, 2)
End Function

'*****************************************************************************
'[概要] セルに適用されるRGBカラーのRedを取得
'[引数] 対象セル(省略時はセル関数の設定されているセル)
'[戻値] Redの値(0～255)
'*****************************************************************************
Public Function Cell2Red(Optional objCell As Range = Nothing) As Long
    Dim TargetCell As Range
    If objCell Is Nothing Then
        Set TargetCell = Application.ThisCell
    Else
        Set TargetCell = objCell
    End If
    
    With TargetCell.Interior
        Select Case .ColorIndex
        Case xlNone, xlAutomatic
            '透明
            Cell2Red = 0
        Case Else
            Dim RGB As TRGB
            Dim Color As TLong
            Color.Long = .Color
            LSet RGB = Color
            Cell2Red = RGB.Red
        End Select
    End With
End Function

'*****************************************************************************
'[概要] セルに適用されるRGBカラーのGreenを取得
'[引数] 対象セル(省略時はセル関数の設定されているセル)
'[戻値] Greenの値(0～255)
'*****************************************************************************
Public Function Cell2Green(Optional objCell As Range = Nothing) As Long
    Dim TargetCell As Range
    If objCell Is Nothing Then
        Set TargetCell = Application.ThisCell
    Else
        Set TargetCell = objCell
    End If
    
    With TargetCell.Interior
        Select Case .ColorIndex
        Case xlNone, xlAutomatic
            '透明
            Cell2Green = 0
        Case Else
            Dim RGB As TRGB
            Dim Color As TLong
            Color.Long = .Color
            LSet RGB = Color
            Cell2Green = RGB.Green
        End Select
    End With
End Function

'*****************************************************************************
'[概要] セルに適用されるRGBカラーのBlueを取得
'[引数] 対象セル(省略時はセル関数の設定されているセル)
'[戻値] Blueの値(0～255)
'*****************************************************************************
Public Function Cell2Blue(Optional objCell As Range = Nothing) As Long
    Dim TargetCell As Range
    If objCell Is Nothing Then
        Set TargetCell = Application.ThisCell
    Else
        Set TargetCell = objCell
    End If
    
    With TargetCell.Interior
        Select Case .ColorIndex
        Case xlNone, xlAutomatic
            '透明
            Cell2Blue = 0
        Case Else
            Dim RGB As TRGB
            Dim Color As TLong
            Color.Long = .Color
            LSet RGB = Color
            Cell2Blue = RGB.Blue
        End Select
    End With
End Function

