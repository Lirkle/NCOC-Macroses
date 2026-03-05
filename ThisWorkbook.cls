VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Const MAX_PER_SLIDE As Long = 60

' Late-binding safe constants (PowerPoint)
Private Const ppLayoutBlank As Long = 12

Private Const msoShapeRectangle As Long = 1
Private Const msoShapeOval As Long = 9

Private Const msoTextOrientationHorizontal As Long = 1

Private Const msoConnectorStraight As Long = 1
Private Const msoConnectorElbow As Long = 2

Private Const msoFalse As Long = 0
Private Const msoTrue As Long = -1

' Sheet1 columns:
' A Name, B Position, C Number, D Location, E Manager
' F Directorate, G Department, H Unit, I Sub-Unit, J Sub-sub-Unit, K Sub-sub-sub-Unit

Public Sub BuildOrgChart_Final_V31()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Sheet1")

    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Dim pptApp As Object, pres As Object
    On Error Resume Next
    Set pptApp = GetObject(, "PowerPoint.Application")
    If pptApp Is Nothing Then Set pptApp = CreateObject("PowerPoint.Application")
    On Error GoTo 0

    pptApp.Visible = True
    Set pres = pptApp.Presentations.Add

    ' Hierarchy columns (F..K)
    Dim levelCols As Variant
    levelCols = Array(6, 7, 8, 9, 10, 11)

    Dim rowsAll As Collection
    Set rowsAll = New Collection

    Dim i As Long
    For i = 2 To lastRow
        If CleanKey(ws.Cells(i, 1).Value) <> "" Then rowsAll.Add CLng(i)
    Next i

    SplitAndDraw_V31 ws, pres, rowsAll, levelCols, 0, ""

End Sub

' =========================================================
' Split rule (strict):
' 1) Try current level group (Directorate first).
' 2) If group size <= 60 -> STOP splitting and draw slides for this group.
' 3) If > 60 -> split by next level (Department, then Unit, etc).
' =========================================================
Private Sub SplitAndDraw_V31(ws As Worksheet, pres As Object, rows As Collection, levelCols As Variant, ByVal level As Long, ByVal path As String)

    If rows.Count = 0 Then Exit Sub

    If rows.Count <= MAX_PER_SLIDE Or level > UBound(levelCols) Then
        DrawManagerSlidesForGroup_V31 ws, pres, rows, path
        Exit Sub
    End If

    Dim groups As Object
    Set groups = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 1 To rows.Count
        Dim key As String
        key = CleanKey(ws.Cells(rows(i), levelCols(level)).Value)
        If key = "" Then key = "(Unassigned)"
        If Not groups.Exists(key) Then groups.Add key, New Collection
        groups(key).Add CLng(rows(i))
    Next i

    Dim g As Variant
    For Each g In groups.Keys
        Dim newPath As String
        If path = "" Then newPath = CStr(g) Else newPath = path & " | " & CStr(g)
        SplitAndDraw_V31 ws, pres, groups(g), levelCols, level + 1, newPath
    Next g

End Sub

' =========================================================
' For a final group (<=60): create slides for each manager with direct reports in this group
' =========================================================
Private Sub DrawManagerSlidesForGroup_V31(ws As Worksheet, pres As Object, rows As Collection, ByVal titleText As String)

    Dim nameSet As Object: Set nameSet = CreateObject("Scripting.Dictionary")
    Dim nameToRow As Object: Set nameToRow = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 1 To rows.Count
        Dim nm As String
        nm = CleanKey(ws.Cells(rows(i), 1).Value)
        If nm <> "" Then
            nameSet(nm) = True
            If Not nameToRow.Exists(nm) Then nameToRow.Add nm, CLng(rows(i))
        End If
    Next i

    Dim children As Object
    Set children = CreateObject("Scripting.Dictionary")

    For i = 1 To rows.Count
        Dim mgr As String
        mgr = CleanKey(ws.Cells(rows(i), 5).Value)
        If mgr <> "" And nameSet.Exists(mgr) Then
            If Not children.Exists(mgr) Then children.Add mgr, New Collection
            children(mgr).Add CLng(rows(i))
        End If
    Next i

    Dim m As Variant
    For Each m In children.Keys
        If nameToRow.Exists(CStr(m)) Then
            DrawManagerSlide_V31 ws, pres, CLng(nameToRow(CStr(m))), children(CStr(m)), titleText
        End If
    Next m

End Sub

' =========================================================
' One slide: Manager + direct reports (two columns only)
' Connectors:
' - 1 report: straight (manager bottom-center -> employee top-center)
' - 2+ reports: elbow forced to L-shape using invisible node (two elbow connectors)
' =========================================================
Private Sub DrawManagerSlide_V31(ws As Worksheet, pres As Object, ByVal mgrRow As Long, reports As Collection, ByVal titleText As String)

    Dim sld As Object
    Set sld = pres.Slides.Add(pres.Slides.Count + 1, ppLayoutBlank)

    Dim SW As Double, SH As Double, cx As Double
    SW = pres.PageSetup.slideWidth
    SH = pres.PageSetup.slideHeight
    cx = SW / 2

    AddTitle_V31 sld, titleText, SW

    Dim repCount As Long: repCount = reports.Count

    Dim col1 As Collection, col2 As Collection
    Set col1 = New Collection
    Set col2 = New Collection

    Dim i As Long
    For i = 1 To repCount
        If i Mod 2 = 1 Then col1.Add reports(i) Else col2.Add reports(i)
    Next i

    Dim rowsMax As Long
    rowsMax = (repCount + 1) \ 2
    If rowsMax < 1 Then rowsMax = 1

    Dim topMargin As Double: topMargin = 40
    Dim gapY As Double: gapY = 24
    Dim gapX As Double: gapX = 190
    Dim boxW As Double: boxW = 260

    Dim availH As Double: availH = SH - topMargin - 40
    Dim boxH As Double
    boxH = (availH / (rowsMax + 2)) * 0.9
    If boxH > 75 Then boxH = 75
    If boxH < 44 Then boxH = 44

    Dim fontBase As Double
    fontBase = boxH * 0.22
    If fontBase < 8 Then fontBase = 8
    If fontBase > 14 Then fontBase = 14

    Dim mgrBox As Object
    Set mgrBox = DrawBoxStyled_V31(sld, ws, mgrRow, cx, topMargin, boxW, boxH, fontBase)

    Dim baseY As Double
    baseY = mgrBox.Top + mgrBox.Height + gapY

    Dim empBox As Object

    ' Column 1 bottom-to-top
    For i = col1.Count To 1 Step -1
        Set empBox = DrawBoxStyled_V31(sld, ws, CLng(col1(i)), cx - gapX, baseY + (i - 1) * (boxH + 14), boxW, boxH, fontBase)
        SafeSetTag_V31 empBox, "col", "1"
        DrawConnSmart_V31 sld, mgrBox, empBox, repCount
    Next i

    ' Column 2 bottom-to-top
    For i = col2.Count To 1 Step -1
        Set empBox = DrawBoxStyled_V31(sld, ws, CLng(col2(i)), cx + gapX, baseY + (i - 1) * (boxH + 14), boxW, boxH, fontBase)
        SafeSetTag_V31 empBox, "col", "2"
        DrawConnSmart_V31 sld, mgrBox, empBox, repCount
    Next i

End Sub

' =========================
' Connectors
' =========================
Private Sub DrawConnSmart_V31(sld As Object, mgrBox As Object, empBox As Object, ByVal totalKids As Long)
    If totalKids <= 1 Then
        AddStraight_V31 sld, mgrBox, empBox
    Else
        AddElbow_EmployeeToManager_V31 sld, mgrBox, empBox
    End If
End Sub

Private Sub AddStraight_V31(sld As Object, fromBox As Object, toBox As Object)

    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
    x1 = fromBox.Left + fromBox.Width / 2
    y1 = fromBox.Top + fromBox.Height
    x2 = toBox.Left + toBox.Width / 2
    y2 = toBox.Top

    Dim c As Object
    Set c = sld.shapes.AddConnector(msoConnectorStraight, x1, y1, x2, y2)

    On Error Resume Next
    c.ConnectorFormat.BeginConnect fromBox, 3
    c.ConnectorFormat.EndConnect toBox, 1
    c.BeginX = x1: c.BeginY = y1
    c.EndX = x2: c.EndY = y2
    c.ConnectorFormat.RerouteConnections
    On Error GoTo 0

    c.Line.Weight = 2
    c.Line.ForeColor.RGB = RGB(0, 112, 192)

End Sub

' Forced L-shaped elbow using an invisible node (two elbow connectors)
Private Sub AddElbow_EmployeeToManager_V31(sld As Object, mgrBox As Object, empBox As Object)

    Dim mgrCX As Double, mgrBY As Double
    mgrCX = mgrBox.Left + mgrBox.Width / 2
    mgrBY = mgrBox.Top + mgrBox.Height

    Dim empCY As Double, empCX As Double
    empCY = empBox.Top + empBox.Height / 2
    empCX = empBox.Left + empBox.Width / 2

    Dim colIndex As Long
    colIndex = 0
    On Error Resume Next
    colIndex = CLng(empBox.Tags("col"))
    On Error GoTo 0
    If colIndex <> 1 And colIndex <> 2 Then colIndex = IIf(empCX < mgrCX, 1, 2)

    Dim empSide As Integer, empX As Double
    If colIndex = 1 Then
        empSide = 2
        empX = empBox.Left + empBox.Width
    Else
        empSide = 4
        empX = empBox.Left
    End If

    Dim kneeX As Double, kneeY As Double
    kneeX = mgrCX
    kneeY = empCY

    Dim node As Object
    Set node = sld.shapes.AddShape(msoShapeOval, kneeX - 1, kneeY - 1, 2, 2)
    node.Line.Visible = msoFalse
    node.Fill.Visible = msoFalse

    Dim c1 As Object
    Set c1 = sld.shapes.AddConnector(msoConnectorElbow, mgrCX, mgrBY, kneeX, kneeY)
    On Error Resume Next
    c1.ConnectorFormat.BeginConnect mgrBox, 3
    c1.ConnectorFormat.EndConnect node, 1
    c1.BeginX = mgrCX: c1.BeginY = mgrBY
    c1.EndX = kneeX:  c1.EndY = kneeY
    On Error GoTo 0
    c1.Line.Weight = 2
    c1.Line.ForeColor.RGB = RGB(0, 112, 192)

    Dim c2 As Object
    Set c2 = sld.shapes.AddConnector(msoConnectorElbow, kneeX, kneeY, empX, empCY)
    On Error Resume Next
    c2.ConnectorFormat.BeginConnect node, 1
    c2.ConnectorFormat.EndConnect empBox, empSide
    c2.BeginX = kneeX: c2.BeginY = kneeY
    c2.EndX = empX:    c2.EndY = empCY
    On Error GoTo 0
    c2.Line.Weight = 2
    c2.Line.ForeColor.RGB = RGB(0, 112, 192)

End Sub

Private Sub SafeSetTag_V31(shp As Object, ByVal key As String, ByVal val As String)
    On Error Resume Next
    shp.Tags.Add key, val
    On Error GoTo 0
End Sub

' =========================
' Boxes
' =========================
Private Function DrawBoxStyled_V31(sld As Object, ws As Worksheet, ByVal r As Long, ByVal cx As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, ByVal f As Double) As Object

    Dim box As Object
    Set box = sld.shapes.AddShape(msoShapeRectangle, cx - w / 2, y, w, h)

    box.Fill.ForeColor.RGB = RGB(255, 255, 255)
    box.Line.ForeColor.RGB = RGB(0, 0, 0)

    With box.TextFrame
        .MarginLeft = 10
        .MarginRight = 10
        .marginTop = 6
        .marginBottom = 6
        .WordWrap = True
    End With

    Dim pos As String, nm As String, num As String, loc As String
    pos = CleanKey(ws.Cells(r, 2).Value)
    nm = CleanKey(ws.Cells(r, 1).Value)
    num = CleanKey(ws.Cells(r, 3).Value)
    loc = CleanKey(ws.Cells(r, 4).Value)

    Dim tr As Object
    Set tr = box.TextFrame.TextRange
    tr.Text = pos & vbCrLf & nm

    tr.Font.name = "Calibri"
    tr.Font.Size = f
    tr.Font.Color.RGB = RGB(0, 0, 0)
    tr.ParagraphFormat.Alignment = 2

    On Error Resume Next
    tr.Paragraphs(1).Font.Color.RGB = RGB(0, 0, 0)
    tr.Paragraphs(2).Font.Color.RGB = RGB(0, 112, 192)
    On Error GoTo 0

    Dim tNum As Object
    Set tNum = sld.shapes.AddTextbox(msoTextOrientationHorizontal, box.Left + 8, box.Top + h - (f + 10), w / 2 - 12, f + 6)
    With tNum.TextFrame.TextRange
        .Text = num
        .Font.name = "Calibri"
        .Font.Size = f - 1
        .Font.Color.RGB = RGB(0, 0, 0)
        .ParagraphFormat.Alignment = 1
    End With
    tNum.Line.Visible = msoFalse
    tNum.Fill.Visible = msoFalse

    Dim tLoc As Object
    Set tLoc = sld.shapes.AddTextbox(msoTextOrientationHorizontal, box.Left + w / 2, box.Top + h - (f + 10), w / 2 - 8, f + 6)
    With tLoc.TextFrame.TextRange
        .Text = loc
        .Font.name = "Calibri"
        .Font.Size = f - 1
        .Font.Color.RGB = RGB(0, 0, 0)
        .ParagraphFormat.Alignment = 3
    End With
    tLoc.Line.Visible = msoFalse
    tLoc.Fill.Visible = msoFalse

    Set DrawBoxStyled_V31 = box

End Function

' =========================
' Title
' =========================
Private Sub AddTitle_V31(sld As Object, ByVal txt As String, ByVal SW As Double)
    Dim t As Object
    Set t = sld.shapes.AddTextbox(msoTextOrientationHorizontal, 20, 10, SW - 40, 24)
    With t.TextFrame.TextRange
        .Text = txt
        .Font.name = "Calibri"
        .Font.Size = 14
        .Font.Color.RGB = RGB(0, 0, 0)
    End With
    t.Line.Visible = msoFalse
    t.Fill.Visible = msoFalse
End Sub

Private Function CleanKey(ByVal v As Variant) As String
    CleanKey = Trim$(Replace(CStr(v), ChrW(160), " "))
End Function

