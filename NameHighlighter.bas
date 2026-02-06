Attribute VB_Name = "NameHighlighter"
Option Explicit

Public Sub HighlightNamesByHexColor()
    Dim hexColor As String
    Dim rgbColor As Long
    Dim reg As Object
    Dim sld As Slide
    Dim shp As Shape
    Dim txtRange As TextRange
    Dim matches As Object
    Dim i As Long

    hexColor = InputBox("Введите HEX-цвет (например, #FFAA00):", "Цвет подсветки")
    If Len(hexColor) = 0 Then
        Exit Sub
    End If

    If Not IsValidHexColor(hexColor) Then
        MsgBox "Некорректный HEX-цвет. Используйте формат #RRGGBB.", vbExclamation
        Exit Sub
    End If

    rgbColor = HexToRgb(hexColor)

    Set reg = CreateObject("VBScript.RegExp")
    reg.Global = True
    reg.IgnoreCase = True
    reg.Pattern = "[А-ЯЁа-яё]{1,2}\.[А-ЯЁа-яё]+"

    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If shp.HasTextFrame Then
                If shp.TextFrame.HasText Then
                    Set txtRange = shp.TextFrame.TextRange
                    Set matches = reg.Execute(txtRange.Text)

                    For i = matches.Count - 1 To 0 Step -1
                        txtRange.Characters(matches(i).FirstIndex + 1, matches(i).Length).Font.Color.RGB = rgbColor
                    Next i
                End If
            End If
        Next shp
    Next sld

    MsgBox "Подсветка завершена.", vbInformation
End Sub

Private Function IsValidHexColor(ByVal value As String) As Boolean
    Dim hexValue As String

    hexValue = Trim$(value)
    If Left$(hexValue, 1) = "#" Then
        hexValue = Mid$(hexValue, 2)
    End If

    If Len(hexValue) <> 6 Then
        IsValidHexColor = False
        Exit Function
    End If

    IsValidHexColor = (hexValue Like "[0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f]")
End Function

Private Function HexToRgb(ByVal value As String) As Long
    Dim hexValue As String
    Dim r As Long
    Dim g As Long
    Dim b As Long

    hexValue = Trim$(value)
    If Left$(hexValue, 1) = "#" Then
        hexValue = Mid$(hexValue, 2)
    End If

    r = CLng("&H" & Mid$(hexValue, 1, 2))
    g = CLng("&H" & Mid$(hexValue, 3, 2))
    b = CLng("&H" & Mid$(hexValue, 5, 2))

    HexToRgb = RGB(r, g, b)
End Function
