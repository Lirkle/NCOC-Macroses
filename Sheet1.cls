Option Explicit

Private Const MAX_PER_SLIDE As Long = 60

' Sheet columns
Private Const COL_NAME As Long = 1
Private Const COL_POSITION As Long = 2
Private Const COL_NUMBER As Long = 3
Private Const COL_LOCATION As Long = 4
Private Const COL_MANAGER As Long = 5
Private Const COL_DIRECTORATE As Long = 6
Private Const COL_DEPARTMENT As Long = 7
Private Const COL_UNIT As Long = 8
Private Const COL_SUBUNIT As Long = 9
Private Const COL_SUBSUBUNIT As Long = 10
Private Const COL_SUBSUBSUBUNIT As Long = 11

' PowerPoint constants for late binding
Private Const ppLayoutBlank As Long = 12
Private Const msoShapeRectangle As Long = 1
Private Const msoTextOrientationHorizontal As Long = 1
Private Const msoConnectorStraight As Long = 1
Private Const msoConnectorElbow As Long = 2
Private Const msoFalse As Long = 0
Private Const msoTrue As Long = -1

' Grid export: one slide keeps up to 60 employees
Public Sub BuildOrgChart_OneSlide60_V31()
    On Error GoTo EH

    Dim ws As Worksheet
    Set ws = Me

    Dim rowsAll As Collection
    Set rowsAll = CollectEmployeeRows(ws)
    If rowsAll.Count = 0 Then
        MsgBox "No employees found in column A.", vbExclamation, "One-Slide Export"
        Exit Sub
    End If

    Dim rowsOneSlide As Collection
    Set rowsOneSlide = SortRowsByName(ws, rowsAll)

    If rowsOneSlide.Count > MAX_PER_SLIDE Then
        Set rowsOneSlide = SliceCollection(rowsOneSlide, 1, MAX_PER_SLIDE)
        MsgBox "Found " & CStr(rowsAll.Count) & " employees." & vbCrLf & _
               "One-slide mode supports up to " & CStr(MAX_PER_SLIDE) & "." & vbCrLf & _
               "Only first " & CStr(MAX_PER_SLIDE) & " employees will be exported.", _
               vbExclamation, "One-Slide Export"
    End If

    Dim pres As Object
    Set pres = CreatePresentation()
    If pres Is Nothing Then
        Err.Raise vbObjectError + 2000, "BuildOrgChart_OneSlide60_V31", "PowerPoint could not be started."
    End If

    DrawOneSlideHierarchy_OneSlide60 ws, pres, rowsOneSlide, "Org chart | Single slide"

    MsgBox "One-slide hierarchy export finished: " & CStr(pres.Slides.Count) & " slide(s).", vbInformation, "One-Slide Export"
    Exit Sub

EH:
    MsgBox "One-slide export failed: " & Err.Description, vbCritical, "One-Slide Export"
End Sub

' Hierarchy export: manager links inside each split group
Public Sub BuildOrgChart_Hierarchy_V34()
    On Error GoTo EH

    Dim ws As Worksheet
    Set ws = Me

    Dim rowsAll As Collection
    Set rowsAll = CollectEmployeeRows(ws)
    If rowsAll.Count = 0 Then
        MsgBox "No employees found in column A.", vbExclamation, "Hierarchy Export"
        Exit Sub
    End If

    Dim pres As Object
    Set pres = CreatePresentation()
    If pres Is Nothing Then
        Err.Raise vbObjectError + 2001, "BuildOrgChart_Hierarchy_V34", "PowerPoint could not be started."
    End If

    Dim levelCols As Variant
    levelCols = Array(COL_DIRECTORATE, COL_DEPARTMENT, COL_UNIT, COL_SUBUNIT, COL_SUBSUBUNIT, COL_SUBSUBSUBUNIT)

    SplitAndDraw_V34 ws, pres, SortRowsByName(ws, rowsAll), levelCols, 0, ""

    MsgBox "Hierarchy export finished: " & CStr(pres.Slides.Count) & " slide(s).", vbInformation, "Hierarchy Export"
    Exit Sub

EH:
    MsgBox "Hierarchy export failed: " & Err.Description, vbCritical, "Hierarchy Export"
End Sub

Private Function CollectEmployeeRows(ws As Worksheet) As Collection
    Dim result As New Collection

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_NAME).End(xlUp).Row

    Dim i As Long
    For i = 2 To lastRow
        If CanonKey(ws.Cells(i, COL_NAME).Value) <> "" Then
            result.Add CLng(i)
        End If
    Next i

    Set CollectEmployeeRows = result
End Function

Private Function CreatePresentation() As Object
    On Error GoTo FailCreate

    Dim pptApp As Object

    On Error Resume Next
    Set pptApp = GetObject(, "PowerPoint.Application")
    On Error GoTo FailCreate

    If pptApp Is Nothing Then
        Set pptApp = CreateObject("PowerPoint.Application")
    End If

    If pptApp Is Nothing Then Exit Function

    pptApp.Visible = True
    Set CreatePresentation = pptApp.Presentations.Add
    Exit Function

FailCreate:
    Set CreatePresentation = Nothing
End Function

Private Sub DrawGridSlide_OneSlide60_V31(ws As Worksheet, pres As Object, rows As Collection, ByVal titleText As String)
    Dim sld As Object
    Set sld = pres.Slides.Add(pres.Slides.Count + 1, ppLayoutBlank)

    Dim SW As Double, SH As Double
    SW = pres.PageSetup.SlideWidth
    SH = pres.PageSetup.SlideHeight

    AddTitle_OneSlide60_V31 sld, titleText, SW, 28

    Dim marginX As Double: marginX = 24
    Dim marginTop As Double: marginTop = 44
    Dim marginBottom As Double: marginBottom = 20
    Dim gapX As Double: gapX = 12
    Dim gapY As Double: gapY = 10

    Dim availW As Double: availW = SW - 2 * marginX
    Dim availH As Double: availH = SH - marginTop - marginBottom

    Dim n As Long
    n = rows.Count
    If n < 1 Then Exit Sub

    Dim cols As Long
    cols = ChooseGridColumns_OneSlide60(n, availW, availH, gapX, gapY)
    If cols < 1 Then cols = 6

    Dim rowsCount As Long
    rowsCount = (n + cols - 1) \ cols

    Dim boxW As Double
    boxW = (availW - (cols - 1) * gapX) / cols
    If boxW > 220 Then boxW = 220

    Dim boxH As Double
    boxH = (availH - (rowsCount - 1) * gapY) / rowsCount
    If boxH > 76 Then boxH = 76
    If boxH < 34 Then boxH = 34

    Dim totalW As Double
    totalW = cols * boxW + (cols - 1) * gapX

    Dim totalH As Double
    totalH = rowsCount * boxH + (rowsCount - 1) * gapY

    If totalW > availW And totalW > 0 Then
        Dim scaleW As Double
        scaleW = availW / totalW
        boxW = boxW * scaleW
        gapX = gapX * scaleW
        totalW = cols * boxW + (cols - 1) * gapX
    End If

    If totalH > availH And totalH > 0 Then
        Dim scaleH As Double
        scaleH = availH / totalH
        boxH = boxH * scaleH
        gapY = gapY * scaleH
        totalH = rowsCount * boxH + (rowsCount - 1) * gapY
    End If

    Dim fontBase As Double
    fontBase = boxH * 0.22
    If fontBase < 7 Then fontBase = 7
    If fontBase > 12 Then fontBase = 12

    Dim startY As Double
    startY = marginTop + (availH - totalH) / 2

    Dim idx As Long
    idx = 1

    Dim rr As Long
    For rr = 0 To rowsCount - 1
        Dim rowItems As Long
        rowItems = n - rr * cols
        If rowItems > cols Then rowItems = cols
        If rowItems < 1 Then Exit For

        Dim rowW As Double
        rowW = rowItems * boxW + (rowItems - 1) * gapX

        Dim rowStartX As Double
        rowStartX = (SW - rowW) / 2 + boxW / 2

        Dim y As Double
        y = startY + rr * (boxH + gapY)

        Dim cc As Long
        For cc = 0 To rowItems - 1
            Dim cx As Double
            cx = rowStartX + cc * (boxW + gapX)

            DrawGridCard_OneSlide60 sld, ws, CLng(rows(idx)), cx, y, boxW, boxH, fontBase
            idx = idx + 1
        Next cc
    Next rr
End Sub

Private Sub AddTitle_OneSlide60_V31(sld As Object, ByVal txt As String, ByVal slideW As Double, ByVal titleH As Double)
    Dim t As Object
    Set t = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 20, 8, slideW - 40, titleH)

    With t.TextFrame.TextRange
        .Text = txt
        .Font.Name = "Calibri"
        .Font.Size = 14
        .Font.Color.RGB = RGB(0, 0, 0)
    End With

    t.Line.Visible = msoFalse
    t.Fill.Visible = msoFalse
End Sub

Private Sub DrawOneSlideHierarchy_OneSlide60(ws As Worksheet, pres As Object, rows As Collection, ByVal titleText As String)
    Dim sld As Object
    Set sld = pres.Slides.Add(pres.Slides.Count + 1, ppLayoutBlank)

    Dim SW As Double, SH As Double
    SW = pres.PageSetup.SlideWidth
    SH = pres.PageSetup.SlideHeight

    AddTitle_OneSlide60_V31 sld, titleText & " (" & CStr(rows.Count) & ")", SW, 28

    Dim nameToId As Object
    Set nameToId = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 1 To rows.Count
        Dim rowId As Long
        rowId = CLng(rows(i))

        Dim nm As String
        nm = CanonKey(ws.Cells(rowId, COL_NAME).Value)
        If nm <> "" Then
            If Not nameToId.Exists(nm) Then nameToId.Add nm, rowId
        End If
    Next i

    Dim parentOf As Object
    Set parentOf = CreateObject("Scripting.Dictionary")

    Dim childrenOf As Object
    Set childrenOf = CreateObject("Scripting.Dictionary")

    For i = 1 To rows.Count
        rowId = CLng(rows(i))

        Dim mgrKey As String
        mgrKey = CanonKey(ws.Cells(rowId, COL_MANAGER).Value)

        If mgrKey <> "" And nameToId.Exists(mgrKey) Then
            Dim mgrId As Long
            mgrId = CLng(nameToId(mgrKey))
            If mgrId <> rowId Then
                parentOf(CStr(rowId)) = mgrId
                If Not childrenOf.Exists(CStr(mgrId)) Then childrenOf.Add CStr(mgrId), New Collection
                childrenOf(CStr(mgrId)).Add rowId
            Else
                parentOf(CStr(rowId)) = 0
            End If
        Else
            parentOf(CStr(rowId)) = 0
        End If
    Next i

    Dim roots As New Collection
    For i = 1 To rows.Count
        rowId = CLng(rows(i))
        If CLng(parentOf(CStr(rowId))) = 0 Then roots.Add rowId
    Next i
    If roots.Count = 0 Then roots.Add CLng(rows(1))

    Dim levelOf As Object
    Set levelOf = CreateObject("Scripting.Dictionary")

    Dim levels As Object
    Set levels = CreateObject("Scripting.Dictionary")

    Dim q As New Collection
    For i = 1 To roots.Count
        rowId = CLng(roots(i))
        levelOf(CStr(rowId)) = 1
        q.Add rowId
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

    For i = 1 To rows.Count
        rowId = CLng(rows(i))
        If Not levelOf.Exists(CStr(rowId)) Then
            levelOf(CStr(rowId)) = 1
            If Not levels.Exists("1") Then levels.Add "1", New Collection
            levels("1").Add rowId
        End If
    Next i

    Dim maxLevel As Long
    maxLevel = 1

    Dim maxN As Long
    maxN = 1

    Dim k As Variant
    For Each k In levels.Keys
        Dim lvKey As Long
        lvKey = CLng(k)
        If lvKey > maxLevel Then maxLevel = lvKey
        If levels(k).Count > maxN Then maxN = levels(k).Count
    Next k

    Dim marginX As Double: marginX = 24
    Dim topY As Double: topY = 46
    Dim bottomMargin As Double: bottomMargin = 20
    Dim gapX As Double: gapX = 12
    Dim gapY As Double: gapY = 20

    Dim availW As Double: availW = SW - 2 * marginX
    Dim availH As Double: availH = SH - topY - bottomMargin

    Dim boxW As Double
    boxW = (availW - (maxN - 1) * gapX) / maxN
    If boxW > 220 Then boxW = 220
    If boxW < 85 Then boxW = 85

    Dim boxH As Double
    boxH = (availH - (maxLevel - 1) * gapY) / maxLevel
    If boxH > 72 Then boxH = 72
    If boxH < 34 Then boxH = 34

    Dim totalW As Double
    totalW = maxN * boxW + (maxN - 1) * gapX
    If totalW > availW And totalW > 0 Then
        Dim sW As Double
        sW = availW / totalW
        boxW = boxW * sW
        gapX = gapX * sW
    End If

    Dim totalH As Double
    totalH = maxLevel * boxH + (maxLevel - 1) * gapY
    If totalH > availH And totalH > 0 Then
        Dim sH As Double
        sH = availH / totalH
        boxH = boxH * sH
        gapY = gapY * sH
    End If

    Dim fontBase As Double
    fontBase = boxH * 0.22
    If fontBase < 7 Then fontBase = 7
    If fontBase > 12 Then fontBase = 12

    Dim shapeOf As Object
    Set shapeOf = CreateObject("Scripting.Dictionary")

    Dim levelIndex As Long
    For levelIndex = 1 To maxLevel
        If Not levels.Exists(CStr(levelIndex)) Then GoTo NextLevelDraw

        Dim nodes As Collection
        Set nodes = SortRowsByName(ws, levels(CStr(levelIndex)))

        Dim countN As Long
        countN = nodes.Count
        If countN < 1 Then GoTo NextLevelDraw

        Dim rowW As Double
        rowW = countN * boxW + (countN - 1) * gapX

        Dim startX As Double
        startX = (SW - rowW) / 2 + boxW / 2

        Dim y As Double
        y = topY + (levelIndex - 1) * (boxH + gapY)

        For i = 1 To countN
            rowId = CLng(nodes(i))

            Dim cx As Double
            cx = startX + (i - 1) * (boxW + gapX)

            Dim shp As Object
            Dim shpKey As String
            shpKey = CStr(rowId)

            Set shp = DrawSimpleBox_OneSlide60(sld, ws, rowId, cx, y, boxW, boxH, fontBase)
            If shapeOf.Exists(shpKey) Then shapeOf.Remove shpKey
            shapeOf.Add shpKey, shp
        Next i

NextLevelDraw:
    Next levelIndex

    Dim parentKeys As Variant
    parentKeys = SortedDictionaryKeys(childrenOf)

    If HasArrayItems(parentKeys) Then
        For i = LBound(parentKeys) To UBound(parentKeys)
            Dim pidKey As String
            pidKey = CStr(parentKeys(i))
            If Not shapeOf.Exists(pidKey) Then GoTo NextParent

            Dim parentShape As Object
            Set parentShape = shapeOf(pidKey)

            Dim childRows As Collection
            Set childRows = SortRowsByName(ws, childrenOf(pidKey))

            Dim childShapes As New Collection
            Dim j As Long
            For j = 1 To childRows.Count
                Dim childId As Long
                childId = CLng(childRows(j))
                If shapeOf.Exists(CStr(childId)) Then
                    Dim childShape As Object
                    Set childShape = shapeOf(CStr(childId))
                    childShapes.Add childShape
                End If
            Next j

            If childShapes.Count = 1 Then
                Dim firstChild As Object
                Set firstChild = childShapes(1)
                DrawSimpleConnector_OneSlide60 sld, parentShape, firstChild
            ElseIf childShapes.Count > 1 Then
                DrawParentBus_OneSlide60 sld, parentShape, childShapes
            End If

NextParent:
        Next i
    End If
End Sub

Private Function DrawSimpleBox_OneSlide60(sld As Object, ws As Worksheet, ByVal rowId As Long, ByVal cx As Double, ByVal y As Double, ByVal boxW As Double, ByVal boxH As Double, ByVal fontSize As Double) As Object
    Dim box As Object
    Set box = sld.Shapes.AddShape(msoShapeRectangle, cx - boxW / 2, y, boxW, boxH)

    box.Fill.ForeColor.RGB = RGB(255, 255, 255)
    box.Line.ForeColor.RGB = RGB(80, 80, 80)
    box.Line.Weight = 0.75

    Dim pos As String
    pos = CleanText(ws.Cells(rowId, COL_POSITION).Value)

    Dim nm As String
    nm = CleanText(ws.Cells(rowId, COL_NAME).Value)

    If pos = "" Then
        box.TextFrame.TextRange.Text = nm
    Else
        box.TextFrame.TextRange.Text = pos & vbCrLf & nm
    End If

    With box.TextFrame.TextRange
        .Font.Name = "Calibri"
        .Font.Size = fontSize
        .ParagraphFormat.Alignment = 2
    End With

    On Error Resume Next
    If pos <> "" Then
        box.TextFrame.TextRange.Paragraphs(1).Font.Color.RGB = RGB(40, 40, 40)
        box.TextFrame.TextRange.Paragraphs(1).Font.Size = fontSize - 1
        box.TextFrame.TextRange.Paragraphs(2).Font.Color.RGB = RGB(0, 112, 192)
        box.TextFrame.TextRange.Paragraphs(2).Font.Bold = msoTrue
    Else
        box.TextFrame.TextRange.Paragraphs(1).Font.Color.RGB = RGB(0, 112, 192)
        box.TextFrame.TextRange.Paragraphs(1).Font.Bold = msoTrue
    End If
    On Error GoTo 0

    Set DrawSimpleBox_OneSlide60 = box
End Function

Private Sub DrawSimpleConnector_OneSlide60(sld As Object, parentBox As Object, childBox As Object)
    Dim px As Double, py As Double, cx As Double, cy As Double
    px = parentBox.Left + parentBox.Width / 2
    py = parentBox.Top + parentBox.Height
    cx = childBox.Left + childBox.Width / 2
    cy = childBox.Top

    Dim midY As Double
    midY = (py + cy) / 2
    If midY <= py + 1 Or midY >= cy - 1 Then
        AddStyledLine_OneSlide60 sld, px, py, cx, cy
    Else
        AddStyledLine_OneSlide60 sld, px, py, px, midY
        AddStyledLine_OneSlide60 sld, px, midY, cx, midY
        AddStyledLine_OneSlide60 sld, cx, midY, cx, cy
    End If
End Sub

Private Sub DrawParentBus_OneSlide60(sld As Object, parentBox As Object, childShapes As Collection)
    Dim parentCX As Double
    parentCX = parentBox.Left + parentBox.Width / 2

    Dim parentBY As Double
    parentBY = parentBox.Top + parentBox.Height

    Dim minChildTop As Double
    minChildTop = 1E+30

    Dim minCX As Double
    minCX = 1E+30

    Dim maxCX As Double
    maxCX = -1E+30

    Dim i As Long
    For i = 1 To childShapes.Count
        Dim childBox As Object
        Set childBox = childShapes(i)

        Dim childCX As Double
        childCX = childBox.Left + childBox.Width / 2

        If childBox.Top < minChildTop Then minChildTop = childBox.Top
        If childCX < minCX Then minCX = childCX
        If childCX > maxCX Then maxCX = childCX
    Next i

    If minChildTop >= 1E+30 Then Exit Sub

    Dim busY As Double
    busY = parentBY + 12
    If busY > minChildTop - 8 Then busY = (parentBY + (minChildTop - 8)) / 2
    If busY <= parentBY + 2 Then busY = parentBY + 2

    If parentCX < minCX Then minCX = parentCX
    If parentCX > maxCX Then maxCX = parentCX

    AddStyledLine_OneSlide60 sld, parentCX, parentBY, parentCX, busY
    AddStyledLine_OneSlide60 sld, minCX, busY, maxCX, busY

    For i = 1 To childShapes.Count
        Set childBox = childShapes(i)
        childCX = childBox.Left + childBox.Width / 2
        AddStyledLine_OneSlide60 sld, childCX, busY, childCX, childBox.Top
    Next i
End Sub

Private Sub AddStyledLine_OneSlide60(sld As Object, ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double)
    If Abs(x1 - x2) < 0.1 And Abs(y1 - y2) < 0.1 Then Exit Sub

    Dim l As Object
    Set l = sld.Shapes.AddLine(x1, y1, x2, y2)
    l.Line.Weight = 1.5
    l.Line.ForeColor.RGB = RGB(0, 112, 192)
End Sub

Private Function ChooseGridColumns_OneSlide60(ByVal n As Long, ByVal availW As Double, ByVal availH As Double, ByVal gapX As Double, ByVal gapY As Double) As Long
    Dim bestCols As Long
    Dim bestScore As Double
    bestScore = -1

    Dim tryCols As Long
    For tryCols = 3 To 12
        Dim tryRows As Long
        tryRows = (n + tryCols - 1) \ tryCols

        Dim tryW As Double
        tryW = (availW - (tryCols - 1) * gapX) / tryCols

        Dim tryH As Double
        tryH = (availH - (tryRows - 1) * gapY) / tryRows

        If tryW > 70 And tryH > 24 Then
            Dim widthFit As Double
            widthFit = tryW / 118
            If widthFit > 1 Then widthFit = 1

            Dim heightFit As Double
            heightFit = tryH / 40
            If heightFit > 1 Then heightFit = 1

            Dim ratioFit As Double
            ratioFit = 1 - Abs((tryW / tryH) - 2.1) / 2.1
            If ratioFit < 0 Then ratioFit = 0

            Dim score As Double
            score = widthFit * 4 + heightFit * 2 + ratioFit

            If score > bestScore Then
                bestScore = score
                bestCols = tryCols
            End If
        End If
    Next tryCols

    If bestCols = 0 Then
        bestCols = n
        If bestCols > 10 Then bestCols = 10
        If bestCols < 2 Then bestCols = 2
    End If

    ChooseGridColumns_OneSlide60 = bestCols
End Function

Private Sub DrawGridCard_OneSlide60(sld As Object, ws As Worksheet, ByVal rowId As Long, ByVal cx As Double, ByVal y As Double, ByVal boxW As Double, ByVal boxH As Double, ByVal fontSize As Double)
    Dim box As Object
    Set box = sld.Shapes.AddShape(msoShapeRectangle, cx - boxW / 2, y, boxW, boxH)

    box.Fill.ForeColor.RGB = RGB(255, 255, 255)
    box.Line.ForeColor.RGB = RGB(140, 140, 140)
    box.Line.Weight = 0.75

    With box.TextFrame
        .MarginLeft = 6
        .MarginRight = 6
        .MarginTop = 4
        .MarginBottom = 4
        .WordWrap = msoTrue
    End With

    Dim pos As String
    pos = CleanText(ws.Cells(rowId, COL_POSITION).Value)

    Dim nm As String
    nm = CleanText(ws.Cells(rowId, COL_NAME).Value)

    Dim num As String
    num = CleanText(ws.Cells(rowId, COL_NUMBER).Value)

    Dim loc As String
    loc = CleanText(ws.Cells(rowId, COL_LOCATION).Value)

    Dim meta As String
    If num <> "" Then meta = num
    If loc <> "" Then
        If meta = "" Then
            meta = loc
        Else
            meta = meta & " | " & loc
        End If
    End If

    Dim hasPos As Boolean
    hasPos = (pos <> "")

    Dim txt As String
    If hasPos Then
        txt = pos & vbCrLf & nm
    Else
        txt = nm
    End If
    If meta <> "" Then txt = txt & vbCrLf & meta

    Dim tr As Object
    Set tr = box.TextFrame.TextRange
    tr.Text = txt
    tr.Font.Name = "Calibri"
    tr.ParagraphFormat.Alignment = 2

    On Error Resume Next
    Dim metaSize As Double
    metaSize = fontSize - 2
    If metaSize < 6.5 Then metaSize = 6.5

    If hasPos Then
        tr.Paragraphs(1).Font.Size = fontSize - 1
        tr.Paragraphs(1).Font.Color.RGB = RGB(40, 40, 40)

        tr.Paragraphs(2).Font.Size = fontSize
        tr.Paragraphs(2).Font.Bold = msoTrue
        tr.Paragraphs(2).Font.Color.RGB = RGB(0, 112, 192)

        If meta <> "" Then
            tr.Paragraphs(3).Font.Size = metaSize
            tr.Paragraphs(3).Font.Color.RGB = RGB(90, 90, 90)
        End If
    Else
        tr.Paragraphs(1).Font.Size = fontSize
        tr.Paragraphs(1).Font.Bold = msoTrue
        tr.Paragraphs(1).Font.Color.RGB = RGB(0, 112, 192)

        If meta <> "" Then
            tr.Paragraphs(2).Font.Size = metaSize
            tr.Paragraphs(2).Font.Color.RGB = RGB(90, 90, 90)
        End If
    End If
    On Error GoTo 0
End Sub

Private Sub SplitAndDraw_V34(ws As Worksheet, pres As Object, rows As Collection, levelCols As Variant, ByVal level As Long, ByVal path As String)
    If rows.Count = 0 Then Exit Sub

    If rows.Count <= MAX_PER_SLIDE Or level > UBound(levelCols) Then
        DrawRowsBatched_V34 ws, pres, rows, path
        Exit Sub
    End If

    Dim splitCol As Long
    splitCol = CLng(levelCols(level))

    Dim groups As Object
    Set groups = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 1 To rows.Count
        Dim rowId As Long
        rowId = CLng(rows(i))

        Dim key As String
        key = CanonKey(ws.Cells(rowId, splitCol).Value)
        If key = "" Then key = "(UNASSIGNED)"

        If Not groups.Exists(key) Then groups.Add key, New Collection
        groups(key).Add rowId
    Next i

    Dim keys As Variant
    keys = SortedDictionaryKeys(groups)
    If Not HasArrayItems(keys) Then
        DrawRowsBatched_V34 ws, pres, rows, path
        Exit Sub
    End If

    If UBound(keys) = LBound(keys) Then
        Dim singlePath As String
        singlePath = AppendPath(path, CStr(keys(LBound(keys))))
        SplitAndDraw_V34 ws, pres, groups(CStr(keys(LBound(keys)))), levelCols, level + 1, singlePath
        Exit Sub
    End If

    For i = LBound(keys) To UBound(keys)
        Dim gKey As String
        gKey = CStr(keys(i))

        Dim newPath As String
        newPath = AppendPath(path, gKey)

        SplitAndDraw_V34 ws, pres, groups(gKey), levelCols, level + 1, newPath
    Next i
End Sub

Private Sub DrawRowsBatched_V34(ws As Worksheet, pres As Object, rows As Collection, ByVal titleText As String)
    Dim sortedRows As Collection
    Set sortedRows = SortRowsByName(ws, rows)

    Dim startIndex As Long
    startIndex = 1

    Do While startIndex <= sortedRows.Count
        Dim endIndex As Long
        endIndex = startIndex + MAX_PER_SLIDE - 1
        If endIndex > sortedRows.Count Then endIndex = sortedRows.Count

        Dim batch As Collection
        Set batch = SliceCollection(sortedRows, startIndex, MAX_PER_SLIDE)

        Dim finalTitle As String
        finalTitle = BuildBatchTitle(titleText, startIndex, endIndex, sortedRows.Count)

        DrawOneSlideOrgChartForGroup_V34 ws, pres, batch, finalTitle

        startIndex = endIndex + 1
    Loop
End Sub

Private Sub DrawOneSlideOrgChartForGroup_V34(ws As Worksheet, pres As Object, rows As Collection, ByVal titleText As String)
    Dim sld As Object
    Set sld = pres.Slides.Add(pres.Slides.Count + 1, ppLayoutBlank)

    Dim SW As Double, SH As Double
    SW = pres.PageSetup.SlideWidth
    SH = pres.PageSetup.SlideHeight

    AddTitle_V34 sld, titleText & " (" & CStr(rows.Count) & ")", SW

    Dim nameToId As Object
    Set nameToId = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 1 To rows.Count
        Dim rowId As Long
        rowId = CLng(rows(i))

        Dim nm As String
        nm = CanonKey(ws.Cells(rowId, COL_NAME).Value)
        If nm <> "" Then
            If Not nameToId.Exists(nm) Then nameToId.Add nm, rowId
        End If
    Next i

    Dim parentOf As Object
    Set parentOf = CreateObject("Scripting.Dictionary")

    Dim childrenOf As Object
    Set childrenOf = CreateObject("Scripting.Dictionary")

    For i = 1 To rows.Count
        rowId = CLng(rows(i))

        Dim mgrKey As String
        mgrKey = CanonKey(ws.Cells(rowId, COL_MANAGER).Value)

        If mgrKey <> "" And nameToId.Exists(mgrKey) Then
            Dim mgrId As Long
            mgrId = CLng(nameToId(mgrKey))

            If mgrId <> rowId Then
                parentOf(CStr(rowId)) = mgrId

                If Not childrenOf.Exists(CStr(mgrId)) Then
                    childrenOf.Add CStr(mgrId), New Collection
                End If
                childrenOf(CStr(mgrId)).Add rowId
            Else
                parentOf(CStr(rowId)) = 0
            End If
        Else
            parentOf(CStr(rowId)) = 0
        End If
    Next i

    Dim roots As New Collection
    For i = 1 To rows.Count
        rowId = CLng(rows(i))
        If CLng(parentOf(CStr(rowId))) = 0 Then roots.Add rowId
    Next i

    If roots.Count = 0 Then roots.Add CLng(rows(1))

    Dim levelOf As Object
    Set levelOf = CreateObject("Scripting.Dictionary")

    Dim levels As Object
    Set levels = CreateObject("Scripting.Dictionary")

    Dim q As New Collection
    For i = 1 To roots.Count
        rowId = CLng(roots(i))
        levelOf(CStr(rowId)) = 1
        q.Add rowId
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

    For i = 1 To rows.Count
        rowId = CLng(rows(i))
        If Not levelOf.Exists(CStr(rowId)) Then
            levelOf(CStr(rowId)) = 1
            If Not levels.Exists("1") Then levels.Add "1", New Collection
            levels("1").Add rowId
        End If
    Next i

    Dim maxLevel As Long
    maxLevel = 1

    Dim maxN As Long
    maxN = 1

    Dim k As Variant
    For Each k In levels.Keys
        Dim lvKey As Long
        lvKey = CLng(k)

        If lvKey > maxLevel Then maxLevel = lvKey
        If levels(k).Count > maxN Then maxN = levels(k).Count
    Next k

    Dim marginX As Double: marginX = 28
    Dim topY As Double: topY = 52
    Dim bottomMargin As Double: bottomMargin = 22
    Dim gapX As Double: gapX = 14
    Dim gapY As Double: gapY = 24

    Dim availW As Double: availW = SW - 2 * marginX
    Dim availH As Double: availH = SH - topY - bottomMargin

    Dim boxW As Double
    boxW = (availW - (maxN - 1) * gapX) / maxN
    If boxW > 250 Then boxW = 250
    If boxW < 80 Then boxW = 80

    Dim boxH As Double
    boxH = (availH - (maxLevel - 1) * gapY) / maxLevel
    If boxH > 74 Then boxH = 74
    If boxH < 30 Then boxH = 30

    Dim totalW As Double
    totalW = maxN * boxW + (maxN - 1) * gapX

    Dim totalH As Double
    totalH = maxLevel * boxH + (maxLevel - 1) * gapY

    If totalW > availW And totalW > 0 Then
        Dim wScale As Double
        wScale = availW / totalW
        boxW = boxW * wScale
        gapX = gapX * wScale
    End If

    If totalH > availH And totalH > 0 Then
        Dim hScale As Double
        hScale = availH / totalH
        boxH = boxH * hScale
        gapY = gapY * hScale
    End If

    Dim fontBase As Double
    fontBase = boxH * 0.22
    If fontBase < 6.5 Then fontBase = 6.5
    If fontBase > 12 Then fontBase = 12

    Dim shapeOf As Object
    Set shapeOf = CreateObject("Scripting.Dictionary")

    Dim levelIndex As Long
    For levelIndex = 1 To maxLevel
        If Not levels.Exists(CStr(levelIndex)) Then GoTo NextLevel

        Dim nodes As Collection
        Set nodes = SortRowsByName(ws, levels(CStr(levelIndex)))

        Dim countN As Long
        countN = nodes.Count

        Dim rowW As Double
        rowW = countN * boxW + (countN - 1) * gapX

        Dim startX As Double
        startX = (SW - rowW) / 2 + boxW / 2

        Dim y As Double
        y = topY + (levelIndex - 1) * (boxH + gapY)

        For i = 1 To countN
            rowId = CLng(nodes(i))

            Dim cx As Double
            cx = startX + (i - 1) * (boxW + gapX)

            Dim shp As Object
            Dim shpKey As String
            shpKey = CStr(rowId)

            Set shp = DrawBoxStyled_V34(sld, ws, rowId, cx, y, boxW, boxH, fontBase)
            If shapeOf.Exists(shpKey) Then shapeOf.Remove shpKey
            shapeOf.Add shpKey, shp
        Next i

NextLevel:
    Next levelIndex

    Dim parentKeys As Variant
    parentKeys = SortedDictionaryKeys(childrenOf)

    If HasArrayItems(parentKeys) Then
        For i = LBound(parentKeys) To UBound(parentKeys)
            Dim pidKey As String
            pidKey = CStr(parentKeys(i))

            If shapeOf.Exists(pidKey) Then
                Dim parentShape As Object
                Set parentShape = shapeOf(pidKey)

                Dim childRows As Collection
                Set childRows = SortRowsByName(ws, childrenOf(pidKey))

                Dim childShapes As New Collection
                Dim j As Long

                For j = 1 To childRows.Count
                    Dim childId As Long
                    childId = CLng(childRows(j))
                    If shapeOf.Exists(CStr(childId)) Then
                        Dim childShape As Object
                        Set childShape = shapeOf(CStr(childId))
                        childShapes.Add childShape
                    End If
                Next j

                If childShapes.Count = 1 Then
                    Dim firstChild As Object
                    Set firstChild = childShapes(1)
                    AddStraightSafe_V34 sld, parentShape, firstChild
                ElseIf childShapes.Count > 1 Then
                    AddBusDownSafe_V34 sld, parentShape, childShapes
                End If
            End If
        Next i
    End If
End Sub

Private Sub AddStraightSafe_V34(sld As Object, parentBox As Object, childBox As Object)
    Dim parentCX As Double
    parentCX = parentBox.Left + parentBox.Width / 2

    Dim parentBY As Double
    parentBY = parentBox.Top + parentBox.Height

    Dim childCX As Double
    childCX = childBox.Left + childBox.Width / 2

    Dim childTY As Double
    childTY = childBox.Top

    Dim midY As Double
    midY = (parentBY + childTY) / 2

    If midY <= parentBY + 1 Or midY >= childTY - 1 Then
        AddStyledLine_V34 sld, parentCX, parentBY, childCX, childTY
        Exit Sub
    End If

    AddStyledLine_V34 sld, parentCX, parentBY, parentCX, midY
    AddStyledLine_V34 sld, parentCX, midY, childCX, midY
    AddStyledLine_V34 sld, childCX, midY, childCX, childTY
End Sub

Private Sub AddElbowSafe_V34(sld As Object, parentBox As Object, childBox As Object)
    Dim parentCX As Double, parentBY As Double
    parentCX = parentBox.Left + parentBox.Width / 2
    parentBY = parentBox.Top + parentBox.Height

    Dim childCY As Double
    childCY = childBox.Top + childBox.Height / 2

    Dim childCX As Double
    childCX = childBox.Left + childBox.Width / 2

    Dim childSide As Long
    Dim childX As Double

    If childCX < parentCX Then
        childSide = 2
        childX = childBox.Left + childBox.Width
    Else
        childSide = 4
        childX = childBox.Left
    End If

    AddStyledLine_V34 sld, parentCX, parentBY, parentCX, childCY
    AddStyledLine_V34 sld, parentCX, childCY, childX, childCY
End Sub

Private Sub AddBusDownSafe_V34(sld As Object, parentBox As Object, childShapes As Collection)
    If childShapes Is Nothing Then Exit Sub
    If childShapes.Count < 2 Then Exit Sub

    Dim parentCX As Double
    parentCX = parentBox.Left + parentBox.Width / 2

    Dim parentBY As Double
    parentBY = parentBox.Top + parentBox.Height

    Dim minChildTop As Double
    minChildTop = 1E+30

    Dim minCX As Double
    minCX = 1E+30

    Dim maxCX As Double
    maxCX = -1E+30

    Dim i As Long
    Dim childBox As Object

    For i = 1 To childShapes.Count
        Set childBox = childShapes(i)

        Dim childCX As Double
        childCX = childBox.Left + childBox.Width / 2

        If childBox.Top < minChildTop Then minChildTop = childBox.Top
        If childCX < minCX Then minCX = childCX
        If childCX > maxCX Then maxCX = childCX
    Next i

    If minChildTop >= 1E+30 Then Exit Sub

    Dim busY As Double
    busY = parentBY + 14

    If busY > minChildTop - 10 Then
        busY = (parentBY + (minChildTop - 10)) / 2
    End If

    If busY <= parentBY + 2 Then busY = parentBY + 2

    If parentCX < minCX Then minCX = parentCX
    If parentCX > maxCX Then maxCX = parentCX

    AddStyledLine_V34 sld, parentCX, parentBY, parentCX, busY
    AddStyledLine_V34 sld, minCX, busY, maxCX, busY

    For i = 1 To childShapes.Count
        Set childBox = childShapes(i)
        childCX = childBox.Left + childBox.Width / 2
        AddStyledLine_V34 sld, childCX, busY, childCX, childBox.Top
    Next i
End Sub

Private Sub AddStyledLine_V34(sld As Object, ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double)
    If Abs(x1 - x2) < 0.1 And Abs(y1 - y2) < 0.1 Then Exit Sub

    Dim l As Object
    Set l = sld.Shapes.AddLine(x1, y1, x2, y2)
    l.Line.Weight = 2
    l.Line.ForeColor.RGB = RGB(0, 112, 192)
End Sub

Private Function DrawBoxStyled_V34(sld As Object, ws As Worksheet, ByVal rowId As Long, ByVal cx As Double, ByVal y As Double, ByVal boxW As Double, ByVal boxH As Double, ByVal fontSize As Double) As Object
    Dim box As Object
    Set box = sld.Shapes.AddShape(msoShapeRectangle, cx - boxW / 2, y, boxW, boxH)

    box.Fill.ForeColor.RGB = RGB(255, 255, 255)
    box.Line.ForeColor.RGB = RGB(0, 0, 0)

    With box.TextFrame
        .MarginLeft = 10
        .MarginRight = 10
        .MarginTop = 6
        .MarginBottom = 6
        .WordWrap = msoTrue
    End With

    Dim pos As String
    pos = CleanText(ws.Cells(rowId, COL_POSITION).Value)

    Dim nm As String
    nm = CleanText(ws.Cells(rowId, COL_NAME).Value)

    Dim num As String
    num = CleanText(ws.Cells(rowId, COL_NUMBER).Value)

    Dim loc As String
    loc = CleanText(ws.Cells(rowId, COL_LOCATION).Value)

    Dim tr As Object
    Set tr = box.TextFrame.TextRange

    tr.Text = pos & vbCrLf & nm
    tr.Font.Name = "Calibri"
    tr.Font.Size = fontSize
    tr.ParagraphFormat.Alignment = 2

    On Error Resume Next
    tr.Paragraphs(1).Font.Color.RGB = RGB(0, 0, 0)
    tr.Paragraphs(2).Font.Color.RGB = RGB(0, 112, 192)
    On Error GoTo 0

    Dim metaTop As Double
    metaTop = box.Top + boxH - (fontSize + 10)

    Dim tNum As Object
    Set tNum = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, box.Left + 8, metaTop, boxW / 2 - 12, fontSize + 6)
    tNum.TextFrame.TextRange.Text = num
    tNum.TextFrame.TextRange.Font.Name = "Calibri"
    tNum.TextFrame.TextRange.Font.Size = fontSize - 1
    tNum.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
    tNum.TextFrame.TextRange.ParagraphFormat.Alignment = 1
    tNum.Line.Visible = msoFalse
    tNum.Fill.Visible = msoFalse

    Dim tLoc As Object
    Set tLoc = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, box.Left + boxW / 2, metaTop, boxW / 2 - 8, fontSize + 6)
    tLoc.TextFrame.TextRange.Text = loc
    tLoc.TextFrame.TextRange.Font.Name = "Calibri"
    tLoc.TextFrame.TextRange.Font.Size = fontSize - 1
    tLoc.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
    tLoc.TextFrame.TextRange.ParagraphFormat.Alignment = 3
    tLoc.Line.Visible = msoFalse
    tLoc.Fill.Visible = msoFalse

    Set DrawBoxStyled_V34 = box
End Function

Private Sub AddTitle_V34(sld As Object, ByVal txt As String, ByVal slideW As Double)
    Dim t As Object
    Set t = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 20, 10, slideW - 40, 28)

    With t.TextFrame.TextRange
        .Text = txt
        .Font.Name = "Calibri"
        .Font.Size = 14
        .Font.Color.RGB = RGB(0, 0, 0)
    End With

    t.Line.Visible = msoFalse
    t.Fill.Visible = msoFalse
End Sub

Private Function BuildBatchTitle(ByVal path As String, ByVal startIndex As Long, ByVal endIndex As Long, ByVal totalCount As Long) As String
    Dim prefix As String
    If path = "" Then
        prefix = "Org chart"
    Else
        prefix = path
    End If

    BuildBatchTitle = prefix & " (" & CStr(startIndex) & "-" & CStr(endIndex) & " of " & CStr(totalCount) & ")"
End Function

Private Function AppendPath(ByVal basePath As String, ByVal nextPart As String) As String
    If basePath = "" Then
        AppendPath = nextPart
    Else
        AppendPath = basePath & " | " & nextPart
    End If
End Function

Private Function SliceCollection(source As Collection, ByVal startIndex As Long, ByVal takeCount As Long) As Collection
    Dim result As New Collection

    If source Is Nothing Then
        Set SliceCollection = result
        Exit Function
    End If

    If startIndex < 1 Then startIndex = 1
    If takeCount < 0 Then takeCount = 0

    Dim lastIndex As Long
    lastIndex = startIndex + takeCount - 1
    If lastIndex > source.Count Then lastIndex = source.Count

    Dim i As Long
    For i = startIndex To lastIndex
        result.Add source(i)
    Next i

    Set SliceCollection = result
End Function

Private Function SortRowsByName(ws As Worksheet, rows As Collection) As Collection
    Dim sorted As New Collection

    Dim i As Long
    For i = 1 To rows.Count
        InsertRowSorted ws, sorted, CLng(rows(i))
    Next i

    Set SortRowsByName = sorted
End Function

Private Sub InsertRowSorted(ws As Worksheet, sorted As Collection, ByVal rowId As Long)
    Dim insertAt As Long
    insertAt = sorted.Count + 1

    Dim i As Long
    For i = 1 To sorted.Count
        If CompareRowsByName(ws, rowId, CLng(sorted(i))) < 0 Then
            insertAt = i
            Exit For
        End If
    Next i

    If insertAt <= sorted.Count Then
        sorted.Add rowId, , insertAt
    Else
        sorted.Add rowId
    End If
End Sub

Private Function CompareRowsByName(ws As Worksheet, ByVal rowA As Long, ByVal rowB As Long) As Long
    Dim nameA As String
    nameA = CanonKey(ws.Cells(rowA, COL_NAME).Value)

    Dim nameB As String
    nameB = CanonKey(ws.Cells(rowB, COL_NAME).Value)

    CompareRowsByName = StrComp(nameA, nameB, vbTextCompare)

    If CompareRowsByName = 0 Then
        If rowA < rowB Then
            CompareRowsByName = -1
        ElseIf rowA > rowB Then
            CompareRowsByName = 1
        End If
    End If
End Function

Private Function SortedDictionaryKeys(dict As Object) As Variant
    If dict Is Nothing Or dict.Count = 0 Then
        SortedDictionaryKeys = Array()
        Exit Function
    End If

    Dim keys() As String
    ReDim keys(0 To dict.Count - 1)

    Dim idx As Long
    idx = 0

    Dim k As Variant
    For Each k In dict.Keys
        keys(idx) = CStr(k)
        idx = idx + 1
    Next k

    If UBound(keys) > LBound(keys) Then
        QuickSortStrings keys, LBound(keys), UBound(keys)
    End If

    SortedDictionaryKeys = keys
End Function

Private Sub QuickSortStrings(ByRef arr() As String, ByVal first As Long, ByVal last As Long)
    Dim i As Long
    Dim j As Long
    Dim pivot As String
    Dim tmp As String

    i = first
    j = last
    pivot = arr((first + last) \ 2)

    Do While i <= j
        Do While StrComp(arr(i), pivot, vbTextCompare) < 0
            i = i + 1
        Loop

        Do While StrComp(arr(j), pivot, vbTextCompare) > 0
            j = j - 1
        Loop

        If i <= j Then
            tmp = arr(i)
            arr(i) = arr(j)
            arr(j) = tmp
            i = i + 1
            j = j - 1
        End If
    Loop

    If first < j Then QuickSortStrings arr, first, j
    If i < last Then QuickSortStrings arr, i, last
End Sub

Private Function HasArrayItems(ByVal arr As Variant) As Boolean
    On Error GoTo NoItems
    HasArrayItems = IsArray(arr) And (UBound(arr) >= LBound(arr))
    Exit Function

NoItems:
    HasArrayItems = False
End Function

Private Function CleanText(ByVal v As Variant) As String
    Dim s As String
    s = CStr(v)

    s = Replace(s, ChrW(160), " ")
    s = Replace(s, vbTab, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")

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
    s = Replace(s, "_", " ")
    s = Replace(s, "(", " ")
    s = Replace(s, ")", " ")
    s = Replace(s, "/", " ")
    s = Replace(s, Chr$(92), " ")
    s = Replace(s, "'", " ")
    s = Replace(s, """", " ")

    s = Replace(s, ChrW(&H2010), " ")
    s = Replace(s, ChrW(&H2011), " ")
    s = Replace(s, ChrW(&H2012), " ")
    s = Replace(s, ChrW(&H2013), " ")
    s = Replace(s, ChrW(&H2014), " ")
    s = Replace(s, ChrW(&H2212), " ")

    s = Trim$(s)

    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    CanonKey = s
End Function

