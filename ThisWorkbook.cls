Option Explicit

Private Const MAX_PER_SLIDE As Long = 60
Private Const MAX_REPORTS_PER_SLIDE As Long = 12

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
Private Const msoShapeOval As Long = 9
Private Const msoTextOrientationHorizontal As Long = 1
Private Const msoConnectorStraight As Long = 1
Private Const msoConnectorElbow As Long = 2
Private Const msoFalse As Long = 0
Private Const msoTrue As Long = -1

Public Sub Export_OrgChart_UI()
    On Error GoTo EH

    Dim mode As Long
    mode = AskExportMode()
    If mode = 0 Then Exit Sub

    If MsgBox(BuildModeSummary(mode), vbOKCancel + vbInformation, "Export Org Chart") <> vbOK Then Exit Sub

    Application.StatusBar = "Export Org Chart: running..."

    Select Case mode
        Case 1
            BuildOrgChart_Final_V31
        Case 2
            Sheet1.BuildOrgChart_Hierarchy_V34
        Case 3
            Sheet1.BuildOrgChart_OneSlide60_V31
        Case Else
            Err.Raise vbObjectError + 1010, "Export_OrgChart_UI", "Unsupported export mode."
    End Select

    Application.StatusBar = False
    Exit Sub

EH:
    Application.StatusBar = False
    MsgBox "Export UI failed: " & Err.Description, vbCritical, "Export Org Chart"
End Sub

Public Sub Export_OrgChart()
    Export_OrgChart_UI
End Sub

Public Sub Export_OrgChart_Help()
    MsgBox BuildModePrompt(), vbInformation, "Export Org Chart - Modes"
End Sub

Public Sub BuildOrgChart_Final_V31()
    On Error GoTo EH

    Dim ws As Worksheet
    Set ws = GetSourceWorksheet()
    If ws Is Nothing Then
        Err.Raise vbObjectError + 1000, "BuildOrgChart_Final_V31", "Source worksheet was not found."
    End If

    Dim rowsAll As Collection
    Set rowsAll = CollectEmployeeRows(ws)
    If rowsAll.Count = 0 Then
        MsgBox "No employees found in column A.", vbExclamation, "Org Chart Export"
        Exit Sub
    End If

    Dim pres As Object
    Set pres = CreatePresentation()
    If pres Is Nothing Then
        Err.Raise vbObjectError + 1001, "BuildOrgChart_Final_V31", "PowerPoint could not be started."
    End If

    Dim levelCols As Variant
    levelCols = Array(COL_DIRECTORATE, COL_DEPARTMENT, COL_UNIT, COL_SUBUNIT, COL_SUBSUBUNIT, COL_SUBSUBSUBUNIT)

    SplitAndDraw_V31 ws, pres, rowsAll, levelCols, 0, ""

    MsgBox "PowerPoint export finished: " & CStr(pres.Slides.Count) & " slide(s).", vbInformation, "Org Chart Export"
    Exit Sub

EH:
    MsgBox "Build failed: " & Err.Description, vbCritical, "Org Chart Export"
End Sub

Private Function AskExportMode() As Long
    Do
        Dim picked As Variant
        picked = Application.InputBox(BuildModePrompt() & vbCrLf & vbCrLf & "Enter mode number (1-3):", "Export Org Chart", 1, Type:=1)

        If VarType(picked) = vbBoolean Then
            AskExportMode = 0
            Exit Function
        End If

        If IsNumeric(picked) Then
            If CLng(picked) >= 1 And CLng(picked) <= 3 Then
                AskExportMode = CLng(picked)
                Exit Function
            End If
        End If

        MsgBox "Please enter 1, 2 or 3 (or press Cancel).", vbExclamation, "Export Org Chart"
    Loop
End Function

Private Function BuildModePrompt() As String
    BuildModePrompt = _
        "Choose export mode:" & vbCrLf & _
        "1. Manager Slides (recommended)" & vbCrLf & _
        "   Split by org columns, then manager + direct reports." & vbCrLf & _
        "2. Hierarchy Tree" & vbCrLf & _
        "   One slide per group, full hierarchy with lines." & vbCrLf & _
        "3. One-Slide Hierarchy (up to 60)" & vbCrLf & _
        "   Full hierarchy with connecting lines on a single slide."
End Function

Private Function BuildModeSummary(ByVal mode As Long) As String
    Select Case mode
        Case 1
            BuildModeSummary = "Mode: Manager Slides" & vbCrLf & _
                               "Best for readable manager-centric reporting."
        Case 2
            BuildModeSummary = "Mode: Hierarchy Tree" & vbCrLf & _
                               "Best for full structural org visualization."
        Case 3
            BuildModeSummary = "Mode: One-Slide Hierarchy (up to 60)" & vbCrLf & _
                               "Best for a compact one-page hierarchy with lines."
        Case Else
            BuildModeSummary = "Mode: Unknown"
    End Select

    BuildModeSummary = BuildModeSummary & vbCrLf & vbCrLf & "Press OK to start export."
End Function

Private Function GetSourceWorksheet() As Worksheet
    On Error Resume Next
    Set GetSourceWorksheet = Sheet1
    On Error GoTo 0

    If GetSourceWorksheet Is Nothing Then
        On Error Resume Next
        Set GetSourceWorksheet = ThisWorkbook.Worksheets("Sheet1")
        On Error GoTo 0
    End If
End Function

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

Private Sub SplitAndDraw_V31(ws As Worksheet, pres As Object, rows As Collection, levelCols As Variant, ByVal level As Long, ByVal path As String)
    If rows.Count = 0 Then Exit Sub

    If rows.Count <= MAX_PER_SLIDE Or level > UBound(levelCols) Then
        DrawManagerSlidesForGroup_V31 ws, pres, rows, path
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

        If Not groups.Exists(key) Then
            groups.Add key, New Collection
        End If

        groups(key).Add rowId
    Next i

    Dim keys As Variant
    keys = SortedDictionaryKeys(groups)
    If Not HasArrayItems(keys) Then
        DrawManagerSlidesForGroup_V31 ws, pres, rows, path
        Exit Sub
    End If

    For i = LBound(keys) To UBound(keys)
        Dim gKey As String
        gKey = CStr(keys(i))

        Dim newPath As String
        newPath = AppendPath(path, gKey)

        SplitAndDraw_V31 ws, pres, groups(gKey), levelCols, level + 1, newPath
    Next i
End Sub

Private Sub DrawManagerSlidesForGroup_V31(ws As Worksheet, pres As Object, rows As Collection, ByVal titleText As String)
    Dim nameToRow As Object
    Set nameToRow = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 1 To rows.Count
        Dim rowId As Long
        rowId = CLng(rows(i))

        Dim empKey As String
        empKey = CanonKey(ws.Cells(rowId, COL_NAME).Value)
        If empKey <> "" Then
            If Not nameToRow.Exists(empKey) Then
                nameToRow.Add empKey, rowId
            End If
        End If
    Next i

    Dim childrenByManager As Object
    Set childrenByManager = CreateObject("Scripting.Dictionary")

    For i = 1 To rows.Count
        rowId = CLng(rows(i))

        Dim mgrKey As String
        mgrKey = CanonKey(ws.Cells(rowId, COL_MANAGER).Value)

        If mgrKey <> "" And nameToRow.Exists(mgrKey) Then
            If Not childrenByManager.Exists(mgrKey) Then
                childrenByManager.Add mgrKey, New Collection
            End If
            childrenByManager(mgrKey).Add rowId
        End If
    Next i

    Dim coveredRows As Object
    Set coveredRows = CreateObject("Scripting.Dictionary")

    Dim managerKeys As Variant
    managerKeys = SortedDictionaryKeys(childrenByManager)

    If HasArrayItems(managerKeys) Then
        For i = LBound(managerKeys) To UBound(managerKeys)
            Dim managerKey As String
            managerKey = CStr(managerKeys(i))

            Dim mgrRow As Long
            mgrRow = CLng(nameToRow(managerKey))
            coveredRows(CStr(mgrRow)) = True

            Dim reportsSorted As Collection
            Set reportsSorted = SortRowsByName(ws, childrenByManager(managerKey))
            MarkRowsCovered coveredRows, reportsSorted

            Dim totalReports As Long
            totalReports = reportsSorted.Count

            Dim pageCount As Long
            pageCount = (totalReports + MAX_REPORTS_PER_SLIDE - 1) \ MAX_REPORTS_PER_SLIDE
            If pageCount < 1 Then pageCount = 1

            Dim pageIndex As Long
            For pageIndex = 1 To pageCount
                Dim startIndex As Long
                startIndex = (pageIndex - 1) * MAX_REPORTS_PER_SLIDE + 1

                Dim pageRows As Collection
                Set pageRows = SliceCollection(reportsSorted, startIndex, MAX_REPORTS_PER_SLIDE)

                DrawManagerSlide_V31 ws, pres, mgrRow, pageRows, titleText, pageIndex, pageCount, totalReports
            Next pageIndex
        Next i
    End If

    Dim unlinked As New Collection
    For i = 1 To rows.Count
        rowId = CLng(rows(i))
        If Not coveredRows.Exists(CStr(rowId)) Then
            unlinked.Add rowId
        End If
    Next i

    If unlinked.Count > 0 Then
        DrawUnlinkedSlides_V31 ws, pres, SortRowsByName(ws, unlinked), titleText
    End If
End Sub

Private Sub DrawManagerSlide_V31(ws As Worksheet, pres As Object, ByVal mgrRow As Long, reports As Collection, ByVal groupTitle As String, ByVal pageIndex As Long, ByVal pageCount As Long, ByVal totalReports As Long)
    Dim sld As Object
    Set sld = pres.Slides.Add(pres.Slides.Count + 1, ppLayoutBlank)

    Dim SW As Double, SH As Double, centerX As Double
    SW = pres.PageSetup.SlideWidth
    SH = pres.PageSetup.SlideHeight
    centerX = SW / 2

    AddTitle_V31 sld, BuildManagerTitle(ws, mgrRow, groupTitle, pageIndex, pageCount, totalReports), SW

    Dim managerTop As Double
    managerTop = 48

    Dim managerW As Double
    managerW = SW - 220
    If managerW > 300 Then managerW = 300
    If managerW < 200 Then managerW = 200

    Dim managerH As Double
    managerH = 74

    Dim managerFont As Double
    managerFont = 12

    Dim mgrBox As Object
    Set mgrBox = DrawBoxStyled_V31(sld, ws, mgrRow, centerX, managerTop, managerW, managerH, managerFont)

    Dim repCount As Long
    repCount = reports.Count
    If repCount = 0 Then Exit Sub

    If repCount = 1 Then
        Dim singleY As Double
        singleY = mgrBox.Top + mgrBox.Height + 26

        Dim singleH As Double
        singleH = SH - singleY - 24
        If singleH > 72 Then singleH = 72
        If singleH < 40 Then singleH = 40

        Dim singleFont As Double
        singleFont = singleH * 0.22
        If singleFont < 8 Then singleFont = 8
        If singleFont > 12 Then singleFont = 12

        Dim singleBox As Object
        Set singleBox = DrawBoxStyled_V31(sld, ws, CLng(reports(1)), centerX, singleY, managerW, singleH, singleFont)
        AddStraight_V31 sld, mgrBox, singleBox
        Exit Sub
    End If

    Dim sideMargin As Double
    sideMargin = 28

    Dim centerGap As Double
    centerGap = 74

    Dim colW As Double
    colW = (SW - 2 * sideMargin - centerGap) / 2
    If colW > 280 Then colW = 280
    If colW < 150 Then
        colW = (SW - 2 * sideMargin - 32) / 2
        If colW < 120 Then colW = 120
        centerGap = SW - 2 * sideMargin - 2 * colW
        If centerGap < 32 Then centerGap = 32
    End If

    Dim leftCX As Double
    leftCX = centerX - (centerGap / 2 + colW / 2)

    Dim rightCX As Double
    rightCX = centerX + (centerGap / 2 + colW / 2)

    Dim leftCount As Long
    leftCount = (repCount + 1) \ 2

    Dim rightCount As Long
    rightCount = repCount - leftCount

    Dim rowsMax As Long
    rowsMax = leftCount
    If rightCount > rowsMax Then rowsMax = rightCount
    If rowsMax < 1 Then rowsMax = 1

    Dim startY As Double
    startY = mgrBox.Top + mgrBox.Height + 22

    Dim bottomMargin As Double
    bottomMargin = 22

    Dim vGap As Double
    vGap = 10

    Dim availH As Double
    availH = SH - startY - bottomMargin

    Dim repH As Double
    If rowsMax = 1 Then
        repH = availH
    Else
        repH = (availH - (rowsMax - 1) * vGap) / rowsMax
    End If

    If repH > 72 Then repH = 72

    If repH < 36 Then
        vGap = 6
        If rowsMax = 1 Then
            repH = availH
        Else
            repH = (availH - (rowsMax - 1) * vGap) / rowsMax
        End If
        If repH < 30 Then repH = 30
    End If

    Dim repFont As Double
    repFont = repH * 0.22
    If repFont < 7 Then repFont = 7
    If repFont > 12 Then repFont = 12

    Dim idx As Long
    Dim rowId As Long
    Dim empBox As Object
    Dim leftBoxes As New Collection
    Dim rightBoxes As New Collection

    For idx = 1 To leftCount
        rowId = CLng(reports(idx))
        Set empBox = DrawBoxStyled_V31(sld, ws, rowId, leftCX, startY + (idx - 1) * (repH + vGap), colW, repH, repFont)
        SafeSetTag_V31 empBox, "col", "1"
        leftBoxes.Add empBox
    Next idx

    For idx = 1 To rightCount
        rowId = CLng(reports(leftCount + idx))
        Set empBox = DrawBoxStyled_V31(sld, ws, rowId, rightCX, startY + (idx - 1) * (repH + vGap), colW, repH, repFont)
        SafeSetTag_V31 empBox, "col", "2"
        rightBoxes.Add empBox
    Next idx

    DrawManagerBus_V31 sld, mgrBox, leftBoxes, rightBoxes
End Sub

Private Sub DrawUnlinkedSlides_V31(ws As Worksheet, pres As Object, rows As Collection, ByVal groupTitle As String)
    Dim startIndex As Long
    startIndex = 1

    Do While startIndex <= rows.Count
        Dim endIndex As Long
        endIndex = startIndex + MAX_PER_SLIDE - 1
        If endIndex > rows.Count Then endIndex = rows.Count

        Dim batch As Collection
        Set batch = SliceCollection(rows, startIndex, MAX_PER_SLIDE)

        Dim titleText As String
        titleText = BuildUnlinkedTitle(groupTitle, startIndex, endIndex, rows.Count)

        DrawUnlinkedGridSlide_V31 ws, pres, batch, titleText

        startIndex = endIndex + 1
    Loop
End Sub

Private Sub DrawUnlinkedGridSlide_V31(ws As Worksheet, pres As Object, rows As Collection, ByVal titleText As String)
    Dim sld As Object
    Set sld = pres.Slides.Add(pres.Slides.Count + 1, ppLayoutBlank)

    Dim SW As Double, SH As Double
    SW = pres.PageSetup.SlideWidth
    SH = pres.PageSetup.SlideHeight

    AddTitle_V31 sld, titleText, SW

    Dim marginX As Double: marginX = 24
    Dim marginTop As Double: marginTop = 44
    Dim marginBottom As Double: marginBottom = 20
    Dim gapX As Double: gapX = 12
    Dim gapY As Double: gapY = 10

    Dim availW As Double: availW = SW - 2 * marginX
    Dim availH As Double: availH = SH - marginTop - marginBottom

    Dim n As Long
    n = rows.Count
    If n = 0 Then Exit Sub

    Dim bestCols As Long
    Dim bestScore As Double
    bestScore = -1

    Dim tryCols As Long
    For tryCols = 2 To 10
        Dim tryRows As Long
        tryRows = (n + tryCols - 1) \ tryCols

        Dim tryW As Double
        tryW = (availW - (tryCols - 1) * gapX) / tryCols

        Dim tryH As Double
        tryH = (availH - (tryRows - 1) * gapY) / tryRows

        If tryW >= 115 And tryH >= 34 Then
            Dim score As Double
            score = tryW + tryH * 1.35
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

    Dim cols As Long
    cols = bestCols

    Dim rowsCount As Long
    rowsCount = (n + cols - 1) \ cols

    Dim boxW As Double
    boxW = (availW - (cols - 1) * gapX) / cols
    If boxW > 240 Then boxW = 240

    Dim boxH As Double
    boxH = (availH - (rowsCount - 1) * gapY) / rowsCount
    If boxH > 74 Then boxH = 74
    If boxH < 30 Then boxH = 30

    Dim totalW As Double
    totalW = cols * boxW + (cols - 1) * gapX

    Dim totalH As Double
    totalH = rowsCount * boxH + (rowsCount - 1) * gapY

    If totalW > availW And totalW > 0 Then
        Dim wScale As Double
        wScale = availW / totalW
        boxW = boxW * wScale
        gapX = gapX * wScale
        totalW = cols * boxW + (cols - 1) * gapX
    End If

    If totalH > availH And totalH > 0 Then
        Dim hScale As Double
        hScale = availH / totalH
        boxH = boxH * hScale
        gapY = gapY * hScale
        totalH = rowsCount * boxH + (rowsCount - 1) * gapY
    End If

    Dim fontBase As Double
    fontBase = boxH * 0.23
    If fontBase < 6.5 Then fontBase = 6.5
    If fontBase > 11 Then fontBase = 11

    Dim startX As Double
    startX = (SW - totalW) / 2 + boxW / 2

    Dim startY As Double
    startY = marginTop + (availH - totalH) / 2

    Dim idx As Long
    For idx = 1 To n
        Dim rr As Long
        rr = (idx - 1) \ cols

        Dim cc As Long
        cc = (idx - 1) Mod cols

        Dim cx As Double
        cx = startX + cc * (boxW + gapX)

        Dim y As Double
        y = startY + rr * (boxH + gapY)

        DrawBoxStyled_V31 sld, ws, CLng(rows(idx)), cx, y, boxW, boxH, fontBase
    Next idx
End Sub

Private Sub DrawManagerBus_V31(sld As Object, mgrBox As Object, leftBoxes As Collection, rightBoxes As Collection)
    Dim mgrCX As Double
    mgrCX = mgrBox.Left + mgrBox.Width / 2

    Dim mgrBY As Double
    mgrBY = mgrBox.Top + mgrBox.Height

    Dim minChildTop As Double
    minChildTop = 1E+30

    Dim i As Long
    Dim b As Object

    For i = 1 To leftBoxes.Count
        Set b = leftBoxes(i)
        If b.Top < minChildTop Then minChildTop = b.Top
    Next i

    For i = 1 To rightBoxes.Count
        Set b = rightBoxes(i)
        If b.Top < minChildTop Then minChildTop = b.Top
    Next i

    If minChildTop >= 1E+30 Then Exit Sub

    Dim busY As Double
    busY = mgrBY + 16

    If busY > minChildTop - 10 Then
        busY = (mgrBY + (minChildTop - 10)) / 2
    End If

    If busY <= mgrBY + 2 Then
        busY = mgrBY + 2
    End If

    AddStyledLine_V31 sld, mgrCX, mgrBY, mgrCX, busY

    If leftBoxes.Count > 0 Then
        DrawSideBus_V31 sld, mgrCX, busY, leftBoxes, True
    End If

    If rightBoxes.Count > 0 Then
        DrawSideBus_V31 sld, mgrCX, busY, rightBoxes, False
    End If
End Sub

Private Sub DrawSideBus_V31(sld As Object, ByVal mgrCX As Double, ByVal busY As Double, boxes As Collection, ByVal isLeft As Boolean)
    If boxes Is Nothing Then Exit Sub
    If boxes.Count = 0 Then Exit Sub

    Dim firstBox As Object
    Set firstBox = boxes(1)

    Dim firstAttachX As Double
    If isLeft Then
        firstAttachX = firstBox.Left + firstBox.Width
    Else
        firstAttachX = firstBox.Left
    End If

    Dim railX As Double
    If isLeft Then
        railX = firstAttachX + 18
    Else
        railX = firstAttachX - 18
    End If

    AddStyledLine_V31 sld, mgrCX, busY, railX, busY

    Dim maxCY As Double
    maxCY = busY

    Dim i As Long
    Dim b As Object
    Dim childCY As Double

    For i = 1 To boxes.Count
        Set b = boxes(i)
        childCY = b.Top + b.Height / 2
        If childCY > maxCY Then maxCY = childCY
    Next i

    AddStyledLine_V31 sld, railX, busY, railX, maxCY

    For i = 1 To boxes.Count
        Set b = boxes(i)
        childCY = b.Top + b.Height / 2

        Dim attachX As Double
        If isLeft Then
            attachX = b.Left + b.Width
        Else
            attachX = b.Left
        End If

        AddStyledLine_V31 sld, railX, childCY, attachX, childCY
    Next i
End Sub

Private Sub AddStyledLine_V31(sld As Object, ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double)
    If Abs(x1 - x2) < 0.1 And Abs(y1 - y2) < 0.1 Then Exit Sub

    Dim l As Object
    Set l = sld.Shapes.AddLine(x1, y1, x2, y2)
    l.Line.Weight = 2
    l.Line.ForeColor.RGB = RGB(0, 112, 192)
End Sub

Private Sub AddStraight_V31(sld As Object, fromBox As Object, toBox As Object)
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
    x1 = fromBox.Left + fromBox.Width / 2
    y1 = fromBox.Top + fromBox.Height
    x2 = toBox.Left + toBox.Width / 2
    y2 = toBox.Top

    AddStyledLine_V31 sld, x1, y1, x2, y2
End Sub

Private Sub AddElbow_EmployeeToManager_V31(sld As Object, mgrBox As Object, empBox As Object)
    Dim mgrCX As Double, mgrBY As Double
    mgrCX = mgrBox.Left + mgrBox.Width / 2
    mgrBY = mgrBox.Top + mgrBox.Height

    Dim empCX As Double, empCY As Double
    empCX = empBox.Left + empBox.Width / 2
    empCY = empBox.Top + empBox.Height / 2

    Dim colIndex As Long
    colIndex = 0

    On Error Resume Next
    colIndex = CLng(empBox.Tags("col"))
    On Error GoTo 0

    If colIndex <> 1 And colIndex <> 2 Then
        If empCX < mgrCX Then
            colIndex = 1
        Else
            colIndex = 2
        End If
    End If

    Dim empSide As Long
    Dim empAttachX As Double

    If colIndex = 1 Then
        empSide = 2
        empAttachX = empBox.Left + empBox.Width
    Else
        empSide = 4
        empAttachX = empBox.Left
    End If

    Dim kneeY As Double
    kneeY = empCY

    AddStyledLine_V31 sld, mgrCX, mgrBY, mgrCX, kneeY
    AddStyledLine_V31 sld, mgrCX, kneeY, empAttachX, empCY
End Sub

Private Sub SafeSetTag_V31(shp As Object, ByVal key As String, ByVal value As String)
    On Error Resume Next
    shp.Tags.Add key, value
    On Error GoTo 0
End Sub

Private Function DrawBoxStyled_V31(sld As Object, ws As Worksheet, ByVal rowId As Long, ByVal cx As Double, ByVal y As Double, ByVal boxW As Double, ByVal boxH As Double, ByVal fontSize As Double) As Object
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

    Set DrawBoxStyled_V31 = box
End Function

Private Sub AddTitle_V31(sld As Object, ByVal txt As String, ByVal slideW As Double)
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

Private Sub MarkRowsCovered(coveredRows As Object, rows As Collection)
    Dim i As Long
    For i = 1 To rows.Count
        coveredRows(CStr(CLng(rows(i)))) = True
    Next i
End Sub

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

Private Function BuildManagerTitle(ws As Worksheet, ByVal mgrRow As Long, ByVal groupTitle As String, ByVal pageIndex As Long, ByVal pageCount As Long, ByVal totalReports As Long) As String
    Dim managerName As String
    managerName = CleanText(ws.Cells(mgrRow, COL_NAME).Value)
    If managerName = "" Then managerName = "(No name)"

    Dim titleText As String
    titleText = managerName & " (" & CStr(totalReports) & " report"
    If totalReports <> 1 Then titleText = titleText & "s"
    titleText = titleText & ")"

    If groupTitle <> "" Then
        titleText = groupTitle & " | " & titleText
    End If

    If pageCount > 1 Then
        titleText = titleText & " | Page " & CStr(pageIndex) & "/" & CStr(pageCount)
    End If

    BuildManagerTitle = titleText
End Function

Private Function BuildUnlinkedTitle(ByVal groupTitle As String, ByVal startIndex As Long, ByVal endIndex As Long, ByVal totalCount As Long) As String
    Dim titleText As String

    If groupTitle <> "" Then
        titleText = groupTitle & " | "
    End If

    titleText = titleText & "Unlinked employees (" & CStr(startIndex) & "-" & CStr(endIndex) & " of " & CStr(totalCount) & ")"
    BuildUnlinkedTitle = titleText
End Function

Private Function AppendPath(ByVal basePath As String, ByVal nextPart As String) As String
    If basePath = "" Then
        AppendPath = nextPart
    Else
        AppendPath = basePath & " | " & nextPart
    End If
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

