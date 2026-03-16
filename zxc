Option Explicit

Private Const OUTPUT_SHEET As String = "Positions"

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
    Dim outputRow As Long
    Dim processedCount As Long
    Dim issueCount As Long
    Dim hasDocs As Boolean
    Dim previousCalculation As XlCalculation
    Dim previousEnableEvents As Boolean
    Dim previousScreenUpdating As Boolean
    Dim previousDisplayAlerts As Boolean

    folderPath = NormalizeFolderPath(folderPath)
    If Len(Dir$(folderPath, vbDirectory)) = 0 Then
        MsgBox "Folder not found: " & folderPath, vbExclamation
        Exit Sub
    End If

    On Error GoTo CleanFail

    previousCalculation = Application.Calculation
    previousEnableEvents = Application.EnableEvents
    previousScreenUpdating = Application.ScreenUpdating
    previousDisplayAlerts = Application.DisplayAlerts

    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.StatusBar = "Preparing workbook..."

    RemoveSheetIfExists "Log"
    Set ws = EnsureSheet(OUTPUT_SHEET)
    PrepareOutputSheet ws

    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False
    wordApp.DisplayAlerts = 0
    wordApp.ScreenUpdating = False

    outputRow = 2

    hasDocs = ProcessWordFiles(folderPath, wordApp, ws, outputRow, processedCount, issueCount)

    If Not hasDocs Then
        MsgBox "No Word files (.docx or .doc) found in: " & folderPath, vbInformation
        GoTo CleanExit
    End If

    FinalizeOutputSheet ws, outputRow - 1

    MsgBox "Done." & vbCrLf & _
           "Processed files: " & processedCount & vbCrLf & _
           "Rows in " & OUTPUT_SHEET & ": " & WorksheetMax(0, outputRow - 2) & vbCrLf & _
           "Files with issues: " & issueCount, vbInformation
    GoTo CleanExit

CleanFail:
    MsgBox "Processing failed: " & Err.Description, vbExclamation

CleanExit:
    On Error Resume Next
    If Not wordApp Is Nothing Then
        wordApp.Quit False
    End If
    Set wordApp = Nothing
    Application.StatusBar = False
    Application.DisplayAlerts = previousDisplayAlerts
    Application.ScreenUpdating = previousScreenUpdating
    Application.EnableEvents = previousEnableEvents
    Application.Calculation = previousCalculation
End Sub

Private Function ProcessWordFiles(ByVal folderPath As String, ByVal wordApp As Object, _
                                  ByVal ws As Worksheet, ByRef outputRow As Long, _
                                  ByRef processedCount As Long, ByRef issueCount As Long) As Boolean
    Dim fileName As String
    Dim filePath As String
    Dim extensionText As String

    fileName = Dir$(folderPath & "\*.*")

    Do While Len(fileName) > 0
        If Left$(fileName, 2) <> "~$" Then
            extensionText = LCase$(Mid$(fileName, InStrRev(fileName, ".")))

            If extensionText = ".docx" Or extensionText = ".doc" Then
                ProcessWordFiles = True
                filePath = folderPath & "\" & fileName
                Application.StatusBar = "Reading " & fileName & "..."
                ProcessOneDocument filePath, wordApp, ws, outputRow, processedCount, issueCount
            End If
        End If

        fileName = Dir$
    Loop
End Function

Private Sub ProcessOneDocument(ByVal filePath As String, ByVal wordApp As Object, ByVal ws As Worksheet, _
                               ByRef outputRow As Long, ByRef processedCount As Long, ByRef issueCount As Long)
    Dim doc As Object
    Dim englishName As String
    Dim russianName As String
    Dim kazakhName As String
    Dim issues As String

    On Error GoTo HandleFail

    Set doc = wordApp.Documents.Open(FileName:=filePath, ConfirmConversions:=False, ReadOnly:=True, _
                                     AddToRecentFiles:=False, Visible:=False)

    ExtractPositionNames doc, englishName, russianName, kazakhName

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
    issueCount = issueCount + 1
    processedCount = processedCount + 1
    Resume CleanExit
End Sub

Private Sub ExtractPositionNames(ByVal doc As Object, ByRef englishName As String, _
                                 ByRef russianName As String, ByRef kazakhName As String)
    Dim paragraph As Object
    Dim textValue As String
    Dim awaitingLanguage As String
    Dim russianLabel As String
    Dim kazakhLabel As String

    russianLabel = RussianPositionLabel()
    kazakhLabel = KazakhPositionLabel()

    For Each paragraph In doc.Paragraphs
        textValue = CleanText(paragraph.Range.Text)
        If Len(textValue) > 0 Then
            If Len(awaitingLanguage) > 0 Then
                Select Case awaitingLanguage
                    Case "en"
                        englishName = textValue
                    Case "ru"
                        russianName = textValue
                    Case "kk"
                        kazakhName = textValue
                End Select

                awaitingLanguage = vbNullString

                If Len(englishName) > 0 And Len(russianName) > 0 And Len(kazakhName) > 0 Then
                    Exit For
                End If
            ElseIf StrComp(textValue, "Position:", vbTextCompare) = 0 Then
                awaitingLanguage = "en"
            ElseIf StrComp(textValue, russianLabel, vbTextCompare) = 0 Then
                awaitingLanguage = "ru"
            ElseIf StrComp(textValue, kazakhLabel, vbTextCompare) = 0 Then
                awaitingLanguage = "kk"
            End If
        End If
    Next paragraph
End Sub

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

Private Sub RemoveSheetIfExists(ByVal sheetName As String)
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If Not ws Is Nothing Then
        ws.Delete
    End If
End Sub

Private Sub PrepareOutputSheet(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Cells(1, 1).Value = "English"
    ws.Cells(1, 2).Value = RussianHeader()
    ws.Cells(1, 3).Value = KazakhHeader()
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

Private Function RussianPositionLabel() As String
    RussianPositionLabel = U(&H414, &H43E, &H43B, &H436, &H43D, &H43E, &H441, &H442, &H44C, &H3A)
End Function

Private Function KazakhPositionLabel() As String
    KazakhPositionLabel = U(&H41B, &H430, &H443, &H430, &H437, &H44B, &H43C, &H20, &H430, &H442, &H430, &H443, &H44B, &H3A)
End Function

Private Function RussianHeader() As String
    RussianHeader = U(&H420, &H443, &H441, &H441, &H43A, &H438, &H439)
End Function

Private Function KazakhHeader() As String
    KazakhHeader = U(&H49A, &H430, &H437, &H430, &H49B, &H448, &H430)
End Function

Private Function U(ParamArray codes() As Variant) As String
    Dim i As Long

    For i = LBound(codes) To UBound(codes)
        U = U & ChrW$(CLng(codes(i)))
    Next i
End Function
