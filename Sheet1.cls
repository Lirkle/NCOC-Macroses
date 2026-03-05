VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Const MAX_PER_SLIDE As Long = 60

' Late-binding safe constants
Private Const ppLayoutBlank As Long = 12
Private Const msoShapeRectangle As Long = 1
Private Const msoTextOrientationHorizontal As Long = 1
Private Const msoConnectorStraight As Long = 1
Private Const msoConnectorElbow As Long = 2
Private Const msoFalse As Long = 0
Private Const msoTrue As Long = -1

' =========================================================
' BUTTON #2 (Grid export): 60 positions per slide (no hierarchy lines)
' Assign this macro to your second button
' =========================================================
Public Sub BuildOrgChart_OneSlide60_V31()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Sheet1")

    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Dim rowsAll As Collection
    Set rowsAll = New Collection

    Dim i As Long
    For i = 2 To lastRow
        If Trim$(CStr(ws.Cells(i, 1).Value)) <> "" Then rowsAll.Add CLng(i)
    Next i
    If rowsAll.Count = 0 Then Exit Sub

    Dim pptApp As Object, pres As Object
    On Error Resume Next
    Set pptApp = GetObject(, "PowerPoint.Application")
    If pptApp Is Nothing Then Set pptApp = CreateObject("PowerPoint.Application")
    On Error GoTo 0

    pptApp.Visible = True
    Set pres = pptApp.Presentations.Add

    Dim startIdx As Long, endIdx As Long
    Dim batch As Collection

    startIdx = 1
    Do While startIdx <= rowsAll.Count
        endIdx = startIdx + MAX_PER_SLIDE - 1
        If endIdx > rowsAll.Count Then endIdx = rowsAll.Count

        Set batch = New Collection
        For i = startIdx To endIdx
            batch.Add CLng(rowsAll(i))
        Next i

        DrawGridSlide_OneSlide60_V31 ws, pres, batch, _
            "Org chart (" & CStr(startIdx) & "-" & CStr(endIdx) & " of " & CStr(rowsAll.Count) & ")"

        startIdx = endIdx + 1
    Loop

    MsgBox "PPT exported (grid mode).", vbInformation

End Sub

Private Sub DrawGridSlide_OneSlide60_V31(ws As Worksheet, pres As Object, rows As Collection, ByVal titleText As String)

    Dim sld As Object
    Set sld = pres.Slides.Add(pres.Slides.Count + 1, ppLayoutBlank)

    Dim SW As Double, SH As Double
    SW = pres.PageSetup.slideWidth
    SH = pres.PageSetup.slideHeight

    Dim titleH As Double: titleH = 28
    AddTitle_OneSlide60_V31 sld, titleText, SW, titleH

    Dim marginX As Double: marginX = 24
    Dim marginTop As Double: marginTop = 10 + titleH + 8
    Dim marginBottom As Double: marginBottom = 20

    Dim availW As Double: availW = SW - 2 * marginX
    Dim availH As Double: availH = SH - marginTop - marginBottom

    Dim n As Long: n = rows.Count
    If n < 1 Then Exit Sub

    Dim gapX As Double: gapX = 14
    Dim gapY As Double: gapY = 12

    Dim bestCols As Long, bestScore As Double
    bestScore = -1

    Dim tryCols As Long
    For tryCols = 4 To 12
        Dim rws As Long
        rws = (n + tryCols - 1) \ tryCols

        Dim boxW As Double, boxH As Double
        boxW = (availW - (tryCols - 1) * gapX) / tryCols
        boxH = (availH - (rws - 1) * gapY) / rws

        If boxW >= 120 And boxH >= 40 Then
            Dim score As Double
            score = boxW + boxH * 1.2
            If score > bestScore Then
                bestScore = score
                bestCols = tryCols
            End If
        End If
    Next tryCols

    If bestCols = 0 Then bestCols = 10

    Dim cols As Long: cols = bestCols
    Dim rowsCount As Long
    rowsCount = (n + cols - 1) \ cols
    If rowsCount < 1 Then rowsCount = 1

    Dim finalW As Double, finalH As Double
    finalW = (availW - (cols - 1) * gapX) / cols
    finalH = (availH - (rowsCount - 1) * gapY) / rowsCount

    If finalW > 260 Then finalW = 260
    If finalH > 78 Then finalH = 78
    If finalH < 40 Then finalH = 40

    Dim f As Double
    f = finalH * 0.23
    If f < 7 Then f = 7
    If f > 12 Then f = 12

    Dim gridW As Double, gridH As Double
    gridW = cols * finalW + (cols - 1) * gapX
    gridH = rowsCount * finalH + (rowsCount - 1) * gapY

    Dim startX As Double, startY As Double
    startX = (SW - gridW) / 2 + finalW / 2
    startY = marginTop + (availH - gridH) / 2

    Dim idx As Long
    For idx = 1 To n
        Dim r As Long
        r = CLng(rows(idx))

        Dim rr As Long, cc As Long
        rr = (idx - 1) \ cols
        cc = (idx - 1) Mod cols

        Dim cx As Double, y As Double
        cx = startX + cc * (finalW + gapX)
        y = startY + rr * (finalH + gapY)

        ' IMPORTANT FIX: use DrawBoxStyled_V34 (exists in your module)
        Dim tmp As Object
        Set tmp = DrawBoxStyled_V34(sld, ws, r, cx, y, finalW, finalH, f)
    Next idx

End Sub

Private Sub AddTitle_OneSlide60_V31(sld As Object, ByVal txt As String, ByVal SW As Double, ByVal titleH As Double)

    Dim t As Object
    Set t = sld.shapes.AddTextbox(msoTextOrientationHorizontal, 20, 8, SW - 40, titleH)

    On Error Resume Next
    With t.TextFrame.TextRange
        .Text = txt
        .Font.name = "Calibri"
        .Font.Size = 14
        .Font.Color.RGB = RGB(0, 0, 0)
    End With
    t.Line.Visible = msoFalse
    t.Fill.Visible = msoFalse
    On Error GoTo 0

End Sub

' =========================================================
' YOUR V34 (one-slide hierarchy) — kept as-is, only formatting cleaned
' =========================================================
Private Sub SplitAndDraw_V34(ws As Worksheet, pres As Object, rows As Collection, levelCols As Variant, ByVal level As Long, ByVal path As String)

    If rows.Count = 0 Then Exit Sub

    If rows.Count <= MAX_PER_SLIDE Or level > UBound(levelCols) Then
        DrawOneSlideOrgChartForGroup_V34 ws, pres, rows, path
        Exit Sub
    End If

    Dim groups As Object
    Set groups = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 1 To rows.Count
        Dim key As String
        key = CanonKey(ws.Cells(rows(i), levelCols(level)).Value)
        If key = "" Then key = "(UNASSIGNED)"
        If Not groups.Exists(key) Then groups.Add key, New Collection
        groups(key).Add CLng(rows(i))
    Next i

    Dim g As Variant
    For Each g In groups.Keys
        Dim newPath As String
        If path = "" Then newPath = CStr(g) Else newPath = path & " | " & CStr(g)
        SplitAndDraw_V34 ws, pres, groups(g), levelCols, level + 1, newPath
    Next g

End Sub

Private Sub DrawOneSlideOrgChartForGroup_V34(ws As Worksheet, pres As Object, rows As Collection, ByVal titleText As String)

    Dim sld As Object
    Set sld = pres.Slides.Add(pres.Slides.Count + 1, ppLayoutBlank)

    Dim SW As Double, SH As Double
    SW = pres.PageSetup.slideWidth
    SH = pres.PageSetup.slideHeight

    AddTitle_V34 sld, titleText & " (" & CStr(rows.Count) & ")", SW

    Dim nameToId As Object: Set nameToId = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 1 To rows.Count
        Dim nm As String
        nm = CanonKey(ws.Cells(rows(i), 1).Value)
        If nm <> "" Then
            If Not nameToId.Exists(nm) Then nameToId.Add nm, CLng(rows(i))
        End If
    Next i

    Dim parentOf As Object: Set parentOf = CreateObject("Scripting.Dictionary")
    Dim childrenOf As Object: Set childrenOf = CreateObject("Scripting.Dictionary")

    For i = 1 To rows.Count
        Dim empId As Long: empId = CLng(rows(i))
        Dim mgrKey As String: mgrKey = CanonKey(ws.Cells(empId, 5).Value)

        If mgrKey <> "" And nameToId.Exists(mgrKey) Then
            Dim mgrId As Long: mgrId = CLng(nameToId(mgrKey))
            parentOf(CStr(empId)) = mgrId

            If Not childrenOf.Exists(CStr(mgrId)) Then childrenOf.Add CStr(mgrId), New Collection
            childrenOf(CStr(mgrId)).Add empId
        Else
            parentOf(CStr(empId)) = 0
        End If
    Next i

    Dim roots As Collection: Set roots = New Collection
    For i = 1 To rows.Count
        Dim rid As Long: rid = CLng(rows(i))
        If CLng(parentOf(CStr(rid))) = 0 Then roots.Add rid
    Next i
    If roots.Count = 0 Then roots.Add CLng(rows(1))

    Dim levelOf As Object: Set levelOf = CreateObject("Scripting.Dictionary")
    Dim levels As Object: Set levels = CreateObject("Scripting.Dictionary")

    Dim q As Collection: Set q = New Collection
    For i = 1 To roots.Count
        levelOf(CStr(roots(i))) = 1
        q.Add roots(i)
    Next i

    Do While q.Count > 0
        Dim cur As Long
        cur = CLng(q(1))
        q.Remove 1

        Dim lv As Long
        lv = CLng(levelOf(CStr(cur)))

        If Not levels.Exists(CStr(lv)) Then levels.Add CStr(lv), New Collection
        levels(CStr(lv)).Add cur

        If childrenOf.Exists(CStr(cur)) Then
            Dim c As Variant
            For Each c In childrenOf(CStr(cur))
                If Not levelOf.Exists(CStr(c)) Then
                    levelOf(CStr(c)) = lv + 1
                    q.Add CLng(c)
                End If
            Next c
        End If
    Loop

    Dim maxLevel As Long: maxLevel = 1
    Dim maxN As Long: maxN = 1

    Dim k As Variant
    For Each k In levels.Keys
        Dim lvKey As Long: lvKey = CLng(k)
        If lvKey > maxLevel Then maxLevel = lvKey
        If levels(k).Count > maxN Then maxN = levels(k).Count
    Next k

    Dim marginX As Double: marginX = 40
    Dim topY As Double: topY = 55
    Dim bottomMargin As Double: bottomMargin = 30
    Dim gapX As Double: gapX = 22
    Dim gapY As Double: gapY = 36

    Dim availW As Double: availW = SW - 2 * marginX
    Dim availH As Double: availH = SH - topY - bottomMargin

    Dim boxW As Double
    boxW = (availW - (maxN - 1) * gapX) / maxN
    If boxW > 260 Then boxW = 260
    If boxW < 150 Then boxW = 150

    Dim boxH As Double
    boxH = (availH - (maxLevel - 1) * gapY) / maxLevel
    If boxH > 75 Then boxH = 75
    If boxH < 42 Then boxH = 42

    Dim fontBase As Double
    fontBase = boxH * 0.22
    If fontBase < 8 Then fontBase = 8
    If fontBase > 14 Then fontBase = 14

    Dim shapeOf As Object: Set shapeOf = CreateObject("Scripting.Dictionary")

    Dim levelIndex As Long
    For levelIndex = 1 To maxLevel
        If Not levels.Exists(CStr(levelIndex)) Then GoTo NextLevel

        Dim nodes As Collection
        Set nodes = levels(CStr(levelIndex))

        Dim countN As Long: countN = nodes.Count
        Dim rowW As Double: rowW = countN * boxW + (countN - 1) * gapX
        Dim startX As Double: startX = (SW - rowW) / 2 + boxW / 2
        Dim y As Double: y = topY + (levelIndex - 1) * (boxH + gapY)

        For i = 1 To countN
            Dim id As Long: id = CLng(nodes(i))
            Dim cx As Double: cx = startX + (i - 1) * (boxW + gapX)

            Dim shp As Object
            Set shp = DrawBoxStyled_V34(sld, ws, id, cx, y, boxW, boxH, fontBase)
            shapeOf(CStr(id)) = shp
        Next i

NextLevel:
    Next levelIndex

    Dim pairs As Collection: Set pairs = New Collection
    For i = 1 To rows.Count
        Dim childId As Long: childId = CLng(rows(i))
        Dim pid As Long: pid = CLng(parentOf(CStr(childId)))
        If pid <> 0 Then
            If shapeOf.Exists(CStr(pid)) And shapeOf.Exists(CStr(childId)) Then
                pairs.Add CStr(pid) & "|" & CStr(childId)
            End If
        End If
    Next i

    Dim idx As Long
    For idx = pairs.Count To 1 Step -1
        Dim parts() As String
        parts = Split(CStr(pairs(idx)), "|")

        Dim pBox As Object: Set pBox = shapeOf(parts(0))
        Dim cBox As Object: Set cBox = shapeOf(parts(1))

        Dim childCount As Long
        childCount = 0
        If childrenOf.Exists(parts(0)) Then childCount = childrenOf(parts(0)).Count

        If childCount <= 1 Then
            AddStraightSafe_V34 sld, pBox, cBox
        Else
            AddElbowSafe_V34 sld, pBox, cBox
        End If
    Next idx

End Sub

Private Sub AddStraightSafe_V34(sld As Object, parentBox As Object, childBox As Object)

    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
    x1 = parentBox.Left + parentBox.Width / 2
    y1 = parentBox.Top + parentBox.Height
    x2 = childBox.Left + childBox.Width / 2
    y2 = childBox.Top

    On Error Resume Next
    Dim c As Object
    Set c = sld.shapes.AddConnector(msoConnectorStraight, x1, y1, x2, y2)

    Err.Clear
    c.ConnectorFormat.BeginConnect parentBox, 3
    If Err.number <> 0 Then GoTo FallbackLine
    Err.Clear
    c.ConnectorFormat.EndConnect childBox, 1
    If Err.number <> 0 Then GoTo FallbackLine

    c.ConnectorFormat.RerouteConnections
    c.Line.Weight = 2
    c.Line.ForeColor.RGB = RGB(0, 112, 192)
    On Error GoTo 0
    Exit Sub

FallbackLine:
    On Error GoTo 0
    Dim l As Object
    Set l = sld.shapes.AddLine(x1, y1, x2, y2)
    l.Line.Weight = 2
    l.Line.ForeColor.RGB = RGB(0, 112, 192)

End Sub

Private Sub AddElbowSafe_V34(sld As Object, parentBox As Object, childBox As Object)

    Dim parentCX As Double, parentBY As Double
    parentCX = parentBox.Left + parentBox.Width / 2
    parentBY = parentBox.Top + parentBox.Height

    Dim childCY As Double
    childCY = childBox.Top + childBox.Height / 2

    Dim childCX As Double
    childCX = childBox.Left + childBox.Width / 2

    Dim childSide As Integer
    Dim childX As Double

    If childCX < parentCX Then
        childSide = 2
        childX = childBox.Left + childBox.Width
    Else
        childSide = 4
        childX = childBox.Left
    End If

    On Error Resume Next
    Dim c As Object
    Set c = sld.shapes.AddConnector(msoConnectorElbow, childX, childCY, parentCX, parentBY)

    Err.Clear
    c.ConnectorFormat.BeginConnect childBox, childSide
    If Err.number <> 0 Then GoTo FallbackElbow
    Err.Clear
    c.ConnectorFormat.EndConnect parentBox, 3
    If Err.number <> 0 Then GoTo FallbackElbow

    c.ConnectorFormat.RerouteConnections
    c.Line.Weight = 2
    c.Line.ForeColor.RGB = RGB(0, 112, 192)
    On Error GoTo 0
    Exit Sub

FallbackElbow:
    On Error GoTo 0
    Dim midX As Double
    midX = (childX + parentCX) / 2

    Dim l1 As Object, l2 As Object
    Set l1 = sld.shapes.AddLine(childX, childCY, midX, childCY)
    Set l2 = sld.shapes.AddLine(midX, childCY, parentCX, parentBY)

    l1.Line.Weight = 2: l1.Line.ForeColor.RGB = RGB(0, 112, 192)
    l2.Line.Weight = 2: l2.Line.ForeColor.RGB = RGB(0, 112, 192)

End Sub

Public Function DrawBoxStyled_V34(sld As Object, ws As Worksheet, ByVal r As Long, ByVal cx As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, ByVal f As Double) As Object

    Dim box As Object
    Set box = sld.shapes.AddShape(msoShapeRectangle, cx - w / 2, y, w, h)

    On Error Resume Next
    box.Fill.ForeColor.RGB = RGB(255, 255, 255)
    box.Line.ForeColor.RGB = RGB(0, 0, 0)
    On Error GoTo 0

    On Error Resume Next
    With box.TextFrame
        .MarginLeft = 10
        .MarginRight = 10
        .marginTop = 6
        .marginBottom = 6
        .WordWrap = msoTrue
    End With
    On Error GoTo 0

    Dim pos As String, nm As String, num As String, loc As String
    pos = CleanText(ws.Cells(r, 2).Value)
    nm = CleanText(ws.Cells(r, 1).Value)
    num = CleanText(ws.Cells(r, 3).Value)
    loc = CleanText(ws.Cells(r, 4).Value)

    On Error Resume Next
    Dim tr As Object
    Set tr = box.TextFrame.TextRange
    tr.Text = pos & vbCrLf & nm
    tr.Font.name = "Calibri"
    tr.Font.Size = f
    tr.ParagraphFormat.Alignment = 2
    tr.Paragraphs(1).Font.Color.RGB = RGB(0, 0, 0)
    tr.Paragraphs(2).Font.Color.RGB = RGB(0, 112, 192)
    On Error GoTo 0

    On Error Resume Next
    Dim tNum As Object
    Set tNum = sld.shapes.AddTextbox(msoTextOrientationHorizontal, box.Left + 8, box.Top + h - (f + 10), w / 2 - 12, f + 6)
    tNum.TextFrame.TextRange.Text = num
    tNum.TextFrame.TextRange.Font.name = "Calibri"
    tNum.TextFrame.TextRange.Font.Size = f - 1
    tNum.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
    tNum.TextFrame.TextRange.ParagraphFormat.Alignment = 1
    tNum.Line.Visible = msoFalse
    tNum.Fill.Visible = msoFalse
    On Error GoTo 0

    On Error Resume Next
    Dim tLoc As Object
    Set tLoc = sld.shapes.AddTextbox(msoTextOrientationHorizontal, box.Left + w / 2, box.Top + h - (f + 10), w / 2 - 8, f + 6)
    tLoc.TextFrame.TextRange.Text = loc
    tLoc.TextFrame.TextRange.Font.name = "Calibri"
    tLoc.TextFrame.TextRange.Font.Size = f - 1
    tLoc.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
    tLoc.TextFrame.TextRange.ParagraphFormat.Alignment = 3
    tLoc.Line.Visible = msoFalse
    tLoc.Fill.Visible = msoFalse
    On Error GoTo 0

    Set DrawBoxStyled_V34 = box

End Function

Private Sub AddTitle_V34(sld As Object, ByVal txt As String, ByVal SW As Double)
    On Error Resume Next
    Dim t As Object
    Set t = sld.shapes.AddTextbox(msoTextOrientationHorizontal, 20, 10, SW - 40, 24)
    t.TextFrame.TextRange.Text = txt
    t.TextFrame.TextRange.Font.name = "Calibri"
    t.TextFrame.TextRange.Font.Size = 14
    t.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
    t.Line.Visible = msoFalse
    t.Fill.Visible = msoFalse
    On Error GoTo 0
End Sub

Private Function CleanText(ByVal v As Variant) As String
    Dim s As String
    s = CStr(v)
    s = Replace(s, ChrW(160), " ")
    s = Replace(s, vbTab, " ")
    s = Trim$(s)
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    CleanText = s
End Function

Private Function CanonKey(ByVal v As Variant) As String
    Dim s As String
    s = UCase$(CleanText(v))
    s = Replace(s, ".", " ")
    s = Replace(s, ",", " ")
    s = Replace(s, "-", " ")
    s = Replace(s, "—", " ")
    s = Replace(s, "–", " ")
    s = Replace(s, "(", " ")
    s = Replace(s, ")", " ")
    s = Replace(s, "/", " ")
    s = Replace(s, "\", " ")
    s = Replace(s, "'", " ")
    s = Replace(s, """", " ")
    s = Trim$(s)
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    CanonKey = s
End Function


