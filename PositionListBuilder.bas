Attribute VB_Name = "PositionListBuilder"
Option Explicit

Private Const OUTPUT_SHEET As String = "Positions"
Private Const LOG_SHEET As String = "Log"
Private Const HEADER_ROW As Long = 1

Public Sub RunPositionImport()
    Dim folderPath As String

    folderPath = PickFolder(ThisWorkbook.Path)
    If Len(folderPath) = 0 Then
        Exit Sub
    End If

    BuildPositionsListFromFolder folderPath
End Sub

Public Sub BuildPositionsListFromFolder(ByVal folderPath As String)
    Dim wordApp As Object
    Dim ws As Worksheet
    Dim logWs As Worksheet
    Dim outputRow As Long
    Dim logRow As Long
    Dim processedCount As Long
    Dim issueCount As Long
    Dim hasDocs As Boolean

    folderPath = NormalizeFolderPath(folderPath)
    If Len(Dir$(folderPath, vbDirectory)) = 0 Then
        MsgBox "Folder not found: " & folderPath, vbExclamation
        Exit Sub
    End If

    On Error GoTo CleanFail

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.StatusBar = "Preparing workbook..."

    Set ws = EnsureSheet(OUTPUT_SHEET)
    Set logWs = EnsureSheet(LOG_SHEET)
    PrepareOutputSheet ws
    PrepareLogSheet logWs

    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False
    wordApp.DisplayAlerts = 0

    outputRow = 2
    logRow = 2

    hasDocs = ProcessPattern(folderPath, "*.docx", wordApp, ws, logWs, outputRow, logRow, processedCount, issueCount)
    hasDocs = ProcessPattern(folderPath, "*.doc", wordApp, ws, logWs, outputRow, logRow, processedCount, issueCount) Or hasDocs

    If Not hasDocs Then
        MsgBox "No Word files (.docx or .doc) found in: " & folderPath, vbInformation
        GoTo CleanExit
    End If

    FinalizeOutputSheet ws, outputRow - 1
    FinalizeLogSheet logWs, logRow - 1

    MsgBox "Done." & vbCrLf & _
           "Processed files: " & processedCount & vbCrLf & _
           "Rows in " & OUTPUT_SHEET & ": " & WorksheetMax(0, outputRow - 2) & vbCrLf & _
           "Issues in " & LOG_SHEET & ": " & WorksheetMax(0, logRow - 2), vbInformation
    GoTo CleanExit

CleanFail:
    LogIssue logWs, logRow, "(runtime)", "Error", Err.Description
    MsgBox "Processing failed: " & Err.Description, vbExclamation

CleanExit:
    On Error Resume Next
    If Not wordApp Is Nothing Then
        wordApp.Quit False
    End If
    Set wordApp = Nothing
    Application.StatusBar = False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

Private Function ProcessPattern(ByVal folderPath As String, ByVal pattern As String, ByVal wordApp As Object, _
                                ByVal ws As Worksheet, ByVal logWs As Worksheet, ByRef outputRow As Long, _
                                ByRef logRow As Long, ByRef processedCount As Long, ByRef issueCount As Long) As Boolean
    Dim fileName As String
    Dim filePath As String

    fileName = Dir$(folderPath & "\" & pattern)

    Do While Len(fileName) > 0
        If Left$(fileName, 2) <> "~$" Then
            ProcessPattern = True
            filePath = folderPath & "\" & fileName
            Application.StatusBar = "Reading " & fileName & "..."
            ProcessOneDocument filePath, wordApp, ws, logWs, outputRow, logRow, processedCount, issueCount
        End If

        fileName = Dir$
    Loop
End Function

Private Sub ProcessOneDocument(ByVal filePath As String, ByVal wordApp As Object, ByVal ws As Worksheet, _
                               ByVal logWs As Worksheet, ByRef outputRow As Long, ByRef logRow As Long, _
                               ByRef processedCount As Long, ByRef issueCount As Long)
    Dim doc As Object
    Dim lines As Collection
    Dim englishName As String
    Dim russianName As String
    Dim kazakhName As String
    Dim issues As String

    On Error GoTo HandleFail

    Set doc = wordApp.Documents.Open(FileName:=filePath, ConfirmConversions:=False, ReadOnly:=True, _
                                     AddToRecentFiles:=False, Visible:=False)

    Set lines = ReadDocumentLines(doc)
    englishName = ValueAfterLabel(lines, "Position:")
    russianName = ValueAfterLabel(lines, "Должность:")
    kazakhName = ValueAfterLabel(lines, "Лауазым атауы:")

    If Len(englishName) = 0 Then issues = AppendIssue(issues, "Missing English title")
    If Len(russianName) = 0 Then issues = AppendIssue(issues, "Missing Russian title")
    If Len(kazakhName) = 0 Then issues = AppendIssue(issues, "Missing Kazakh title")

    If Len(englishName) > 0 Or Len(russianName) > 0 Or Len(kazakhName) > 0 Then
        ws.Cells(outputRow, 1).Value = englishName
        ws.Cells(outputRow, 2).Value = russianName
        ws.Cells(outputRow, 3).Value = kazakhName
        ws.Cells(outputRow, 4).Value = Mid$(filePath, InStrRev(filePath, "\") + 1)
        outputRow = outputRow + 1
    End If

    If Len(issues) > 0 Then
        LogIssue logWs, logRow, Mid$(filePath, InStrRev(filePath, "\") + 1), "Warning", issues
        issueCount = issueCount + 1
    End If

    processedCount = processedCount + 1

CleanExit:
    On Error Resume Next
    If Not doc Is Nothing Then
        doc.Close False
    End If
    Set doc = Nothing
    Exit Sub

HandleFail:
    LogIssue logWs, logRow, Mid$(filePath, InStrRev(filePath, "\") + 1), "Error", Err.Description
    issueCount = issueCount + 1
    processedCount = processedCount + 1
    Resume CleanExit
End Sub

Private Function ReadDocumentLines(ByVal doc As Object) As Collection
    Dim lines As New Collection
    Dim paragraph As Object
    Dim textValue As String

    For Each paragraph In doc.Paragraphs
        textValue = CleanText(paragraph.Range.Text)
        If Len(textValue) > 0 Then
            lines.Add textValue
        End If
    Next paragraph

    Set ReadDocumentLines = lines
End Function

Private Function ValueAfterLabel(ByVal lines As Collection, ByVal labelText As String) As String
    Dim i As Long

    For i = 1 To lines.Count - 1
        If StrComp(lines(i), labelText, vbTextCompare) = 0 Then
            ValueAfterLabel = lines(i + 1)
            Exit Function
        End If
    Next i
End Function

Private Function CleanText(ByVal sourceText As String) As String
    Dim result As String

    result = sourceText
    result = Replace(result, Chr$(13), " ")
    result = Replace(result, Chr$(11), " ")
    result = Replace(result, Chr$(7), " ")
    result = Replace(result, vbTab, " ")
    result = Replace(result, ChrW$(160), " ")

    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop

    CleanText = Trim$(result)
End Function

Private Function EnsureSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set EnsureSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If EnsureSheet Is Nothing Then
        Set EnsureSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        EnsureSheet.Name = sheetName
    End If
End Function

Private Sub PrepareOutputSheet(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Cells(1, 1).Value = "English"
    ws.Cells(1, 2).Value = "Русский"
    ws.Cells(1, 3).Value = "Қазақша"
    ws.Cells(1, 4).Value = "Source file"
End Sub

Private Sub FinalizeOutputSheet(ByVal ws As Worksheet, ByVal lastRow As Long)
    Dim targetLastRow As Long

    targetLastRow = WorksheetMax(2, lastRow)

    ws.Range("A1:D1").Font.Bold = True
    ws.Range("A1:D" & targetLastRow).WrapText = True
    ws.Columns("A").ColumnWidth = 34
    ws.Columns("B").ColumnWidth = 52
    ws.Columns("C").ColumnWidth = 52
    ws.Columns("D").ColumnWidth = 36
    ws.Rows(1).AutoFilter
    ws.Activate
    ActiveWindow.SplitRow = 1
    ActiveWindow.FreezePanes = True
End Sub

Private Sub PrepareLogSheet(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Cells(1, 1).Value = "File"
    ws.Cells(1, 2).Value = "Status"
    ws.Cells(1, 3).Value = "Details"
End Sub

Private Sub FinalizeLogSheet(ByVal ws As Worksheet, ByVal lastRow As Long)
    Dim targetLastRow As Long

    targetLastRow = WorksheetMax(2, lastRow)

    ws.Range("A1:C1").Font.Bold = True
    ws.Range("A1:C" & targetLastRow).WrapText = True
    ws.Columns("A").ColumnWidth = 38
    ws.Columns("B").ColumnWidth = 14
    ws.Columns("C").ColumnWidth = 80
    ws.Rows(1).AutoFilter
End Sub

Private Sub LogIssue(ByVal ws As Worksheet, ByRef logRow As Long, ByVal fileName As String, _
                     ByVal statusText As String, ByVal detailsText As String)
    ws.Cells(logRow, 1).Value = fileName
    ws.Cells(logRow, 2).Value = statusText
    ws.Cells(logRow, 3).Value = detailsText
    logRow = logRow + 1
End Sub

Private Function AppendIssue(ByVal existingText As String, ByVal newText As String) As String
    If Len(existingText) = 0 Then
        AppendIssue = newText
    Else
        AppendIssue = existingText & "; " & newText
    End If
End Function

Private Function PickFolder(ByVal initialFolder As String) As String
    With Application.FileDialog(4)
        .Title = "Select folder with Word documents"
        .AllowMultiSelect = False
        If Len(initialFolder) > 0 Then
            .InitialFileName = NormalizeFolderPath(initialFolder) & "\"
        End If

        If .Show = -1 Then
            PickFolder = .SelectedItems(1)
        End If
    End With
End Function

Private Function NormalizeFolderPath(ByVal folderPath As String) As String
    NormalizeFolderPath = Trim$(folderPath)
    Do While Right$(NormalizeFolderPath, 1) = "\"
        NormalizeFolderPath = Left$(NormalizeFolderPath, Len(NormalizeFolderPath) - 1)
    Loop
End Function

Private Function WorksheetMax(ByVal a As Long, ByVal b As Long) As Long
    If a > b Then
        WorksheetMax = a
    Else
        WorksheetMax = b
    End If
End Function
