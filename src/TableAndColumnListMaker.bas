Attribute VB_Name = "TableAndColumnListMaker"
Option Explicit

Private Const ERROR_VALUE As String = "ERROR"
Private Const SEARCH_FORMULA_LAST_ROW As Long = 1000
Private Const OUTPUT_FLUSH_ROW_COUNT As Long = 100
Private Const COLUMN_READ_CHUNK_ROW_COUNT As Long = 500
Private Const FORMAT_AUTOFIT_SAMPLE_ROW_COUNT As Long = 2000
Private Const SCAN_YIELD_INTERVAL As Long = 100
Private Const WORKBOOK_YIELD_INTERVAL As Long = 10
Private Const CHECKPOINT_WORKBOOK_INTERVAL As Long = 25
Private Const CHECKPOINT_COLUMN_FLUSH_INTERVAL As Long = 10

Private Type RunStats
    ScannedFileCount As Long
    MatchedFileCount As Long
    ProcessedWorkbookCount As Long
    SkippedWorkbookCount As Long
    SkippedFolderCount As Long
    CheckpointSaveCount As Long
    CheckpointSaveFailureCount As Long
    TableCount As Long
    ColumnCount As Long
End Type

Public Sub RunTableAndColumnListMaker()
    Dim targetFolder As String
    targetFolder = PickTargetFolder()
    If Len(targetFolder) = 0 Then
        Exit Sub
    End If

    RunTableAndColumnListMakerForFolder targetFolder
End Sub

Public Sub RunTableAndColumnListMakerForFolder(ByVal targetFolder As String, Optional ByVal showCompletionMessage As Boolean = True)

    If Len(Trim$(targetFolder)) = 0 Then
        Exit Sub
    End If

    Dim outputBook As Workbook
    Set outputBook = ThisWorkbook

    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldStatusBar As Variant
    Dim oldAskToUpdateLinks As Boolean
    Dim oldAutomationSecurity As MsoAutomationSecurity
    Dim oldCalculation As XlCalculation

    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldDisplayAlerts = Application.DisplayAlerts
    oldStatusBar = Application.StatusBar
    oldAskToUpdateLinks = Application.AskToUpdateLinks
    oldAutomationSecurity = Application.AutomationSecurity
    oldCalculation = Application.Calculation

    On Error GoTo FatalError

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    Application.AutomationSecurity = msoAutomationSecurityForceDisable
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Preparing analysis..."

    Dim tableListSheet As Worksheet
    Dim columnListSheet As Worksheet
    Dim tableSearchSheet As Worksheet
    Dim columnSearchSheet As Worksheet

    Set tableListSheet = ResetWorksheet(outputBook, SheetTableListName())
    Set columnListSheet = ResetWorksheet(outputBook, SheetColumnListName())
    Set tableSearchSheet = GetOrCreateWorksheet(outputBook, SheetTableSearchName())
    Set columnSearchSheet = GetOrCreateWorksheet(outputBook, SheetColumnSearchName())

    SetupTableListSheet tableListSheet
    SetupColumnListSheet columnListSheet

    Dim nextTableRow As Long
    Dim nextColumnRow As Long
    nextTableRow = 2
    nextColumnRow = 2

    Dim stats As RunStats
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(targetFolder) Then
        Err.Raise vbObjectError + 100, "RunTableAndColumnListMaker", "Folder not found: " & targetFolder
    End If

    Dim targetFiles As Collection
    Set targetFiles = CollectTargetWorkbookFiles(targetFolder, stats)

    ProcessWorkbookFiles targetFiles, outputBook, tableListSheet, columnListSheet, nextTableRow, nextColumnRow, stats

    stats.TableCount = nextTableRow - 2
    stats.ColumnCount = nextColumnRow - 2

    SetupTableSearchSheet tableSearchSheet, nextTableRow - 1
    SetupColumnSearchSheet columnSearchSheet, nextColumnRow - 1
    FormatOutputSheet tableListSheet
    FormatOutputSheet columnListSheet
    FormatOutputSheet tableSearchSheet
    FormatOutputSheet columnSearchSheet

CleanExit:
    Application.StatusBar = oldStatusBar
    Application.Calculation = oldCalculation
    Application.AutomationSecurity = oldAutomationSecurity
    Application.AskToUpdateLinks = oldAskToUpdateLinks
    Application.DisplayAlerts = oldDisplayAlerts
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating

    If Err.Number = 0 Then
        CheckpointOutputWorkbook outputBook, stats, "Saving final output..."
        Application.StatusBar = oldStatusBar
    End If

    If Err.Number = 0 And showCompletionMessage Then
        MsgBox BuildCompletionMessage(stats), vbInformation, "TableAndColumnListMaker"
    End If
    Exit Sub

FatalError:
    Dim message As String
    message = "Analysis was interrupted." & vbCrLf & vbCrLf & Err.Description
    Resume RestoreAfterFatalError

RestoreAfterFatalError:
    Application.StatusBar = oldStatusBar
    Application.Calculation = oldCalculation
    Application.AutomationSecurity = oldAutomationSecurity
    Application.AskToUpdateLinks = oldAskToUpdateLinks
    Application.DisplayAlerts = oldDisplayAlerts
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating
    If showCompletionMessage Then
        MsgBox message, vbCritical, "TableAndColumnListMaker"
    Else
        On Error GoTo 0
        Err.Raise vbObjectError + 900, "RunTableAndColumnListMakerForFolder", message
    End If
End Sub

Private Function PickTargetFolder() As String
    Dim dialog As FileDialog
    Set dialog = Application.FileDialog(msoFileDialogFolderPicker)

    With dialog
        .Title = "Select target folder"
        .AllowMultiSelect = False
        If .Show <> -1 Then
            PickTargetFolder = vbNullString
        Else
            PickTargetFolder = CStr(.SelectedItems(1))
        End If
    End With
End Function

Private Function CollectTargetWorkbookFiles(ByVal rootFolderPath As String, ByRef stats As RunStats) As Collection
    Dim targetFiles As New Collection
    Dim folders As New Collection
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    folders.Add rootFolderPath

    Do While folders.Count > 0
        Dim folderPath As String
        folderPath = CStr(folders(1))
        folders.Remove 1

        Dim folder As Object
        On Error Resume Next
        Set folder = fso.GetFolder(folderPath)
        If Err.Number <> 0 Then
            stats.SkippedFolderCount = stats.SkippedFolderCount + 1
            Err.Clear
            On Error GoTo 0
            GoTo ContinueFolderLoop
        End If
        On Error GoTo 0

        Dim file As Object
        On Error Resume Next
        For Each file In folder.Files
            If Err.Number <> 0 Then
                stats.SkippedFolderCount = stats.SkippedFolderCount + 1
                Err.Clear
                Exit For
            End If

            stats.ScannedFileCount = stats.ScannedFileCount + 1
            If IsTargetWorkbookFile(CStr(file.Path)) Then
                stats.MatchedFileCount = stats.MatchedFileCount + 1
                targetFiles.Add CStr(file.Path)
            End If
        Next file
        On Error GoTo 0

        Dim subFolder As Object
        On Error Resume Next
        For Each subFolder In folder.SubFolders
            If Err.Number <> 0 Then
                stats.SkippedFolderCount = stats.SkippedFolderCount + 1
                Err.Clear
                Exit For
            End If
            folders.Add CStr(subFolder.Path)
        Next subFolder
        On Error GoTo 0

ContinueFolderLoop:
        If stats.ScannedFileCount Mod SCAN_YIELD_INTERVAL = 0 Then
            Application.StatusBar = "Scanning files: " & CStr(stats.ScannedFileCount)
            DoEvents
        End If
    Loop

    Set CollectTargetWorkbookFiles = targetFiles
End Function

Private Sub ProcessWorkbookFiles( _
    ByVal targetFiles As Collection, _
    ByVal outputBook As Workbook, _
    ByVal tableListSheet As Worksheet, _
    ByVal columnListSheet As Worksheet, _
    ByRef nextTableRow As Long, _
    ByRef nextColumnRow As Long, _
    ByRef stats As RunStats)

    Dim tableRows As Collection
    Dim columnRows As Collection
    Set tableRows = New Collection
    Set columnRows = New Collection

    Dim index As Long
    For index = 1 To targetFiles.Count
        Application.StatusBar = "Analyzing " & CStr(index) & " / " & CStr(targetFiles.Count) & ": " & CStr(targetFiles(index))
        ProcessWorkbook CStr(targetFiles(index)), outputBook, tableRows, columnRows, tableListSheet, columnListSheet, nextTableRow, nextColumnRow, stats

        FlushBufferedRows tableRows, tableListSheet, 3, nextTableRow
        FlushBufferedRows columnRows, columnListSheet, 4, nextColumnRow

        If index Mod CHECKPOINT_WORKBOOK_INTERVAL = 0 Then
            CheckpointOutputWorkbook outputBook, stats, "Saving progress: " & CStr(index) & " / " & CStr(targetFiles.Count)
        End If

        If index Mod WORKBOOK_YIELD_INTERVAL = 0 Then
            Application.CutCopyMode = False
            DoEvents
        End If
    Next index

    FlushBufferedRows tableRows, tableListSheet, 3, nextTableRow
    FlushBufferedRows columnRows, columnListSheet, 4, nextColumnRow
End Sub

Private Sub ProcessWorkbook( _
    ByVal workbookPath As String, _
    ByVal outputBook As Workbook, _
    ByRef tableRows As Collection, _
    ByRef columnRows As Collection, _
    ByVal tableListSheet As Worksheet, _
    ByVal columnListSheet As Worksheet, _
    ByRef nextTableRow As Long, _
    ByRef nextColumnRow As Long, _
    ByRef stats As RunStats)

    Dim sourceBook As Workbook
    Dim sourceSheet As Worksheet

    On Error GoTo WorkbookError
    Set sourceBook = Workbooks.Open( _
        Filename:=workbookPath, _
        UpdateLinks:=0, _
        ReadOnly:=True, _
        AddToMru:=False, _
        Password:="", _
        WriteResPassword:="", _
        IgnoreReadOnlyRecommended:=True, _
        Notify:=False, _
        CorruptLoad:=xlRepairFile)

    If sourceBook.Worksheets.Count = 0 Then
        Err.Raise vbObjectError + 101, "ProcessWorkbook", "No worksheet found."
    End If

    Set sourceSheet = sourceBook.Worksheets(1)

    Dim englishTableName As String
    englishTableName = ResolveEnglishTableName(sourceSheet)

    Dim japaneseTableName As String
    Dim tableId As String
    ParseSheetName sourceSheet.Name, japaneseTableName, tableId

    CollectColumnRows sourceSheet, outputBook, columnRows, columnListSheet, englishTableName, nextColumnRow, stats

    AddBufferedRow tableRows, englishTableName, japaneseTableName, tableId
    FlushBufferedRowsIfNeeded tableRows, tableListSheet, 3, nextTableRow

    stats.ProcessedWorkbookCount = stats.ProcessedWorkbookCount + 1

CleanWorkbook:
    On Error Resume Next
    Set sourceSheet = Nothing
    If Not sourceBook Is Nothing Then
        sourceBook.Close SaveChanges:=False
    End If
    Set sourceBook = Nothing
    On Error GoTo 0
    Exit Sub

WorkbookError:
    stats.SkippedWorkbookCount = stats.SkippedWorkbookCount + 1
    Debug.Print "Skipped: " & workbookPath & " / " & Err.Description
    Resume CleanWorkbook
End Sub

Private Function IsTargetWorkbookFile(ByVal filePath As String) As Boolean
    Dim fileName As String
    fileName = Mid$(filePath, InStrRev(filePath, "\") + 1)

    If Left$(fileName, 2) = "~$" Then
        Exit Function
    End If

    Dim extension As String
    extension = LCase$(Mid$(fileName, InStrRev(fileName, ".") + 1))
    If extension <> "xlsx" And extension <> "xlsm" And extension <> "xlsb" And extension <> "xls" Then
        Exit Function
    End If

    Dim baseName As String
    baseName = Left$(fileName, Len(fileName) - Len(extension) - 1)
    IsTargetWorkbookFile = IsTargetBaseName(baseName)
End Function

Private Function IsTargetBaseName(ByVal baseName As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^((?:[1-5]|7)-\d{1,3}|[1-5](?!-))(?=$|[^0-9A-Za-z])"
    re.Global = False
    re.IgnoreCase = False
    IsTargetBaseName = re.Test(baseName)
End Function

Private Function ResolveEnglishTableName(ByVal sourceSheet As Worksheet) As String
    Dim c1Value As String
    Dim b2Value As String
    c1Value = TrimRangeText(sourceSheet.Range("C1"))
    b2Value = TrimRangeText(sourceSheet.Range("B2"))

    If Len(c1Value) = 0 And Len(b2Value) = 0 Then
        ResolveEnglishTableName = UnresolvedValue()
    ElseIf Len(c1Value) >= Len(b2Value) Then
        ResolveEnglishTableName = c1Value
    Else
        ResolveEnglishTableName = b2Value
    End If
End Function

Private Sub CollectColumnRows( _
    ByVal sourceSheet As Worksheet, _
    ByVal outputBook As Workbook, _
    ByRef columnRows As Collection, _
    ByVal columnListSheet As Worksheet, _
    ByVal englishTableName As String, _
    ByRef nextColumnRow As Long, _
    ByRef stats As RunStats)

    Dim fieldHeaderCell As Range
    Set fieldHeaderCell = FindFieldHeaderCell(sourceSheet)
    If fieldHeaderCell Is Nothing Then
        Exit Sub
    End If

    Dim fieldColumn As Long
    Dim japaneseColumn As Long
    Dim headerRow As Long
    fieldColumn = fieldHeaderCell.Column
    japaneseColumn = fieldColumn + 1
    headerRow = fieldHeaderCell.Row

    Dim lastRow As Long
    lastRow = LastNonEmptyRowInColumns(sourceSheet, fieldColumn, japaneseColumn)
    If lastRow <= headerRow Then
        Exit Sub
    End If

    Dim chunkStartRow As Long
    Dim chunkEndRow As Long
    Dim values As Variant
    Dim rowIndex As Long
    Dim fieldName As String
    Dim japaneseName As String
    Dim chunkFlushCount As Long

    For chunkStartRow = headerRow + 1 To lastRow Step COLUMN_READ_CHUNK_ROW_COUNT
        chunkEndRow = chunkStartRow + COLUMN_READ_CHUNK_ROW_COUNT - 1
        If chunkEndRow > lastRow Then
            chunkEndRow = lastRow
        End If

        values = sourceSheet.Range(sourceSheet.Cells(chunkStartRow, fieldColumn), sourceSheet.Cells(chunkEndRow, japaneseColumn)).Value2

        For rowIndex = 1 To UBound(values, 1)
            fieldName = TrimCellText(values(rowIndex, 1))
            If Len(fieldName) > 0 Then
                japaneseName = TrimCellText(values(rowIndex, 2))
                AddBufferedRow columnRows, englishTableName, fieldName, japaneseName, BuildColumnLookupKey(englishTableName, fieldName)
            End If
        Next rowIndex

        values = Empty
        If FlushBufferedRowsIfNeeded(columnRows, columnListSheet, 4, nextColumnRow) Then
            chunkFlushCount = chunkFlushCount + 1
            If chunkFlushCount Mod CHECKPOINT_COLUMN_FLUSH_INTERVAL = 0 Then
                CheckpointOutputWorkbook outputBook, stats, "Saving progress inside workbook..."
            End If
        End If
    Next chunkStartRow
End Sub

Private Function FindFieldHeaderCell(ByVal sourceSheet As Worksheet) As Range
    Dim searchRange As Range
    Set searchRange = sourceSheet.Cells

    Dim foundCell As Range
    Set foundCell = searchRange.Find( _
        What:=FieldNameHeaderText(), _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False)

    If foundCell Is Nothing Then
        Exit Function
    End If

    Dim firstAddress As String
    firstAddress = foundCell.Address

    Do
        If foundCell.Column < sourceSheet.Columns.Count Then
            If TrimCellText(foundCell.Offset(0, 1).Value) = JapaneseNameHeaderText() Then
                Set FindFieldHeaderCell = foundCell
                Exit Function
            End If
        End If

        Set foundCell = searchRange.FindNext(foundCell)
        If foundCell Is Nothing Then
            Exit Function
        End If
    Loop While foundCell.Address <> firstAddress
End Function

Private Function LastNonEmptyRowInColumns(ByVal sourceSheet As Worksheet, ByVal firstColumn As Long, ByVal secondColumn As Long) As Long
    Dim firstLastRow As Long
    Dim secondLastRow As Long

    firstLastRow = LastNonEmptyRowInColumn(sourceSheet, firstColumn)
    secondLastRow = LastNonEmptyRowInColumn(sourceSheet, secondColumn)

    If firstLastRow > secondLastRow Then
        LastNonEmptyRowInColumns = firstLastRow
    Else
        LastNonEmptyRowInColumns = secondLastRow
    End If
End Function

Private Function LastNonEmptyRowInColumn(ByVal sourceSheet As Worksheet, ByVal columnIndex As Long) As Long
    Dim foundCell As Range
    Set foundCell = sourceSheet.Columns(columnIndex).Find( _
        What:="*", _
        LookIn:=xlValues, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False)

    If foundCell Is Nothing Then
        LastNonEmptyRowInColumn = 0
    Else
        LastNonEmptyRowInColumn = foundCell.Row
    End If
End Function

Private Sub ParseSheetName(ByVal sheetName As String, ByRef japaneseName As String, ByRef tableId As String)
    Dim tokens As Collection
    Set tokens = TokenizeSheetName(sheetName)

    Dim japaneseParts As New Collection
    Dim token As Variant

    For Each token In tokens
        If ContainsJapanese(CStr(token)) Then
            japaneseParts.Add CStr(token)
        End If
    Next token

    japaneseName = JoinCollection(japaneseParts, " ")
    If Len(japaneseName) = 0 Then
        japaneseName = UnresolvedValue()
    End If

    tableId = ExtractFirstDigitSequence(sheetName)
    If Len(tableId) = 0 Then
        tableId = UnresolvedValue()
    End If
End Sub

Private Function ExtractFirstDigitSequence(ByVal value As String) As String
    Dim index As Long
    Dim started As Boolean
    Dim currentChar As String

    For index = 1 To Len(value)
        currentChar = Mid$(value, index, 1)

        If IsDigitChar(currentChar) Then
            ExtractFirstDigitSequence = ExtractFirstDigitSequence & NormalizeDigitChar(currentChar)
            started = True
        ElseIf started Then
            Exit Function
        End If
    Next index
End Function

Private Function TokenizeSheetName(ByVal sheetName As String) As Collection
    Dim tokens As New Collection
    Dim currentToken As String
    Dim index As Long

    For index = 1 To Len(sheetName)
        Dim currentChar As String
        Dim nextChar As String
        currentChar = Mid$(sheetName, index, 1)
        If index < Len(sheetName) Then
            nextChar = Mid$(sheetName, index + 1, 1)
        Else
            nextChar = vbNullString
        End If

        If IsSheetNameSeparator(currentChar) Then
            AddToken tokens, currentToken
        ElseIf currentChar = "_" Then
            If ContainsJapanese(currentToken) Or IsJapaneseChar(nextChar) Then
                AddToken tokens, currentToken
            Else
                currentToken = currentToken & currentChar
            End If
        ElseIf IsJapaneseChar(currentChar) Then
            currentToken = currentToken & currentChar
        ElseIf IsAsciiAlphaNumeric(currentChar) Then
            If Len(currentToken) > 0 And ContainsJapanese(currentToken) Then
                AddToken tokens, currentToken
            End If
            currentToken = currentToken & currentChar
        Else
            AddToken tokens, currentToken
        End If
    Next index

    AddToken tokens, currentToken
    Set TokenizeSheetName = tokens
End Function

Private Sub AddToken(ByVal tokens As Collection, ByRef currentToken As String)
    currentToken = Trim$(currentToken)
    If Len(currentToken) > 0 Then
        tokens.Add currentToken
    End If
    currentToken = vbNullString
End Sub

Private Function IsSheetNameSeparator(ByVal value As String) As Boolean
    If Len(value) = 0 Then
        Exit Function
    End If

    Dim codePoint As Long
    codePoint = AscW(Left$(value, 1))
    If codePoint < 0 Then
        codePoint = codePoint + 65536
    End If

    Select Case codePoint
        Case 9, 32, 40, 41, 44, 45, 47, 58, 59, 91, 92, 93, _
             12288, 12300, 12301, 65288, 65289, 65292, 65293, _
             65306, 65307, 65339, 65341, 65295, 65340, 12540
            IsSheetNameSeparator = True
    End Select
End Function

Private Function ContainsJapanese(ByVal value As String) As Boolean
    Dim index As Long
    For index = 1 To Len(value)
        If IsJapaneseChar(Mid$(value, index, 1)) Then
            ContainsJapanese = True
            Exit Function
        End If
    Next index
End Function

Private Function IsJapaneseChar(ByVal value As String) As Boolean
    If Len(value) = 0 Then
        Exit Function
    End If

    Dim codePoint As Long
    codePoint = AscW(Left$(value, 1))
    If codePoint < 0 Then
        codePoint = codePoint + 65536
    End If

    IsJapaneseChar = _
        (codePoint >= 12352 And codePoint <= 12447) Or _
        (codePoint >= 12448 And codePoint <= 12543) Or _
        (codePoint >= 19968 And codePoint <= 40959) Or _
        (codePoint >= 63744 And codePoint <= 64255) Or _
        (codePoint >= 65382 And codePoint <= 65437) Or _
        codePoint = 12293 Or _
        codePoint = 12294 Or _
        codePoint = 12295
End Function

Private Function IsAsciiAlphaNumeric(ByVal value As String) As Boolean
    If Len(value) = 0 Then
        Exit Function
    End If

    Dim codePoint As Long
    codePoint = AscW(Left$(value, 1))
    IsAsciiAlphaNumeric = _
        (codePoint >= AscW("0") And codePoint <= AscW("9")) Or _
        (codePoint >= AscW("A") And codePoint <= AscW("Z")) Or _
        (codePoint >= AscW("a") And codePoint <= AscW("z"))
End Function

Private Function IsDigitChar(ByVal value As String) As Boolean
    If Len(value) = 0 Then
        Exit Function
    End If

    Dim codePoint As Long
    codePoint = AscW(Left$(value, 1))
    If codePoint < 0 Then
        codePoint = codePoint + 65536
    End If

    IsDigitChar = _
        (codePoint >= AscW("0") And codePoint <= AscW("9")) Or _
        (codePoint >= 65296 And codePoint <= 65305)
End Function

Private Function NormalizeDigitChar(ByVal value As String) As String
    Dim codePoint As Long
    codePoint = AscW(Left$(value, 1))
    If codePoint < 0 Then
        codePoint = codePoint + 65536
    End If

    If codePoint >= 65296 And codePoint <= 65305 Then
        NormalizeDigitChar = Chr$(codePoint - 65296 + AscW("0"))
    Else
        NormalizeDigitChar = Left$(value, 1)
    End If
End Function

Private Function TrimCellText(ByVal value As Variant) As String
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then
        TrimCellText = vbNullString
    Else
        TrimCellText = Trim$(CStr(value))
    End If
End Function

Private Function TrimRangeText(ByVal targetRange As Range) As String
    If targetRange.MergeCells Then
        TrimRangeText = TrimCellText(targetRange.MergeArea.Cells(1, 1).Value)
    Else
        TrimRangeText = TrimCellText(targetRange.Value)
    End If
End Function

Private Function JoinCollection(ByVal values As Collection, ByVal delimiter As String) As String
    Dim result As String
    Dim index As Long

    For index = 1 To values.Count
        If Len(result) > 0 Then
            result = result & delimiter
        End If
        result = result & CStr(values(index))
    Next index

    JoinCollection = result
End Function

Private Function ResetWorksheet(ByVal targetBook As Workbook, ByVal sheetName As String) As Worksheet
    Dim sheet As Worksheet
    Set sheet = GetOrCreateWorksheet(targetBook, sheetName)
    sheet.Cells.Clear
    Set ResetWorksheet = sheet
End Function

Private Function GetOrCreateWorksheet(ByVal targetBook As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateWorksheet = targetBook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateWorksheet Is Nothing Then
        Set GetOrCreateWorksheet = targetBook.Worksheets.Add(After:=targetBook.Worksheets(targetBook.Worksheets.Count))
        GetOrCreateWorksheet.Name = sheetName
    End If
End Function

Private Sub SetupTableListSheet(ByVal sheet As Worksheet)
    sheet.Range("A1").Value = HeaderTableEnglish()
    sheet.Range("B1").Value = HeaderTableJapanese()
    sheet.Range("C1").Value = HeaderTableId()
End Sub

Private Sub SetupColumnListSheet(ByVal sheet As Worksheet)
    sheet.Range("A1").Value = HeaderTableEnglish()
    sheet.Range("B1").Value = HeaderColumnEnglish()
    sheet.Range("C1").Value = HeaderColumnJapanese()
    sheet.Range("D1").Value = "__KEY"
    sheet.Columns("D").Hidden = True
End Sub

Private Sub SetupTableSearchSheet(ByVal sheet As Worksheet, ByVal lastListRow As Long)
    sheet.Range("A1").Value = HeaderTableEnglish()
    sheet.Range("B1").Value = HeaderTableJapanese()
    sheet.Range("C1").Value = HeaderTableId()
    sheet.Range("B2:C" & SEARCH_FORMULA_LAST_ROW).ClearContents
    sheet.Range("B2:B" & SEARCH_FORMULA_LAST_ROW).Formula = BuildTableSearchFormula("B", lastListRow)
    sheet.Range("C2:C" & SEARCH_FORMULA_LAST_ROW).Formula = BuildTableSearchFormula("C", lastListRow)
End Sub

Private Sub SetupColumnSearchSheet(ByVal sheet As Worksheet, ByVal lastListRow As Long)
    sheet.Range("A1").Value = HeaderTableEnglish()
    sheet.Range("B1").Value = HeaderColumnEnglish()
    sheet.Range("C1").Value = HeaderColumnJapanese()
    sheet.Range("C2:C" & SEARCH_FORMULA_LAST_ROW).ClearContents
    sheet.Range("C2:C" & SEARCH_FORMULA_LAST_ROW).Formula = BuildColumnSearchFormula(lastListRow)
End Sub

Private Function BuildTableSearchFormula(ByVal returnColumn As String, ByVal lastListRow As Long) As String
    If lastListRow < 2 Then
        lastListRow = 2
    End If

    BuildTableSearchFormula = _
        "=IF(TRIM($A2)="""",""""," & _
        "IFERROR(INDEX('" & SheetTableListName() & "'!$" & returnColumn & "$2:$" & returnColumn & "$" & lastListRow & "," & _
        "MATCH(TRIM($A2),'" & SheetTableListName() & "'!$A$2:$A$" & lastListRow & ",0)),""" & ERROR_VALUE & """))"
End Function

Private Function BuildColumnSearchFormula(ByVal lastListRow As Long) As String
    If lastListRow < 2 Then
        lastListRow = 2
    End If

    BuildColumnSearchFormula = _
        "=IF(OR(TRIM($A2)="""",TRIM($B2)=""""),""""," & _
        "IFERROR(INDEX('" & SheetColumnListName() & "'!$C$2:$C$" & lastListRow & "," & _
        "MATCH(TRIM($A2)&""|""&TRIM($B2),'" & SheetColumnListName() & "'!$D$2:$D$" & lastListRow & ",0)),""" & ERROR_VALUE & """))"
End Function

Private Function BuildColumnLookupKey(ByVal englishTableName As String, ByVal englishColumnName As String) As String
    BuildColumnLookupKey = Trim$(englishTableName) & "|" & Trim$(englishColumnName)
End Function

Private Sub AddBufferedRow(ByVal rowBuffer As Collection, ParamArray values() As Variant)
    Dim rowValues As Variant
    rowValues = values
    rowBuffer.Add rowValues
End Sub

Private Sub FlushBufferedRows( _
    ByRef rowBuffer As Collection, _
    ByVal targetSheet As Worksheet, _
    ByVal columnCount As Long, _
    ByRef nextRow As Long)

    If rowBuffer.Count = 0 Then
        Exit Sub
    End If

    If nextRow + rowBuffer.Count - 1 > targetSheet.Rows.Count Then
        Err.Raise vbObjectError + 102, "FlushBufferedRows", "The output sheet row limit was exceeded: " & targetSheet.Name
    End If

    Dim data() As Variant
    ReDim data(1 To rowBuffer.Count, 1 To columnCount)

    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim rowValues As Variant

    For rowIndex = 1 To rowBuffer.Count
        rowValues = rowBuffer(rowIndex)
        For columnIndex = 1 To columnCount
            data(rowIndex, columnIndex) = rowValues(columnIndex - 1)
        Next columnIndex
    Next rowIndex

    targetSheet.Cells(nextRow, 1).Resize(rowBuffer.Count, columnCount).Value2 = data
    nextRow = nextRow + rowBuffer.Count
    Set rowBuffer = New Collection
End Sub

Private Function FlushBufferedRowsIfNeeded( _
    ByRef rowBuffer As Collection, _
    ByVal targetSheet As Worksheet, _
    ByVal columnCount As Long, _
    ByRef nextRow As Long) As Boolean

    If rowBuffer.Count >= OUTPUT_FLUSH_ROW_COUNT Then
        FlushBufferedRows rowBuffer, targetSheet, columnCount, nextRow
        FlushBufferedRowsIfNeeded = True
    End If
End Function

Private Sub CheckpointOutputWorkbook( _
    ByVal outputBook As Workbook, _
    ByRef stats As RunStats, _
    ByVal statusText As String)

    If Len(outputBook.Path) = 0 Then
        Exit Sub
    End If

    If outputBook.ReadOnly Then
        Exit Sub
    End If

    On Error GoTo SaveError
    Application.StatusBar = statusText
    outputBook.Save
    stats.CheckpointSaveCount = stats.CheckpointSaveCount + 1
    DoEvents
    Exit Sub

SaveError:
    stats.CheckpointSaveFailureCount = stats.CheckpointSaveFailureCount + 1
    Debug.Print "Checkpoint save failed: " & Err.Description
    Err.Clear
End Sub

Private Function SheetTableListName() As String
    SheetTableListName = TextFromCodePoints("12486 12540 12502 12523 19968 35239")
End Function

Private Function SheetColumnListName() As String
    SheetColumnListName = TextFromCodePoints("12459 12521 12512 19968 35239")
End Function

Private Function SheetTableSearchName() As String
    SheetTableSearchName = TextFromCodePoints("12486 12540 12502 12523 26908 32034")
End Function

Private Function SheetColumnSearchName() As String
    SheetColumnSearchName = TextFromCodePoints("12459 12521 12512 26908 32034")
End Function

Private Function UnresolvedValue() As String
    UnresolvedValue = TextFromCodePoints("26410 35299 27770")
End Function

Private Function HeaderTableEnglish() As String
    HeaderTableEnglish = TextFromCodePoints("12486 12540 12502 12523 21517 65288 33521 65289")
End Function

Private Function HeaderTableJapanese() As String
    HeaderTableJapanese = TextFromCodePoints("12486 12540 12502 12523 21517 65288 21644 65289")
End Function

Private Function HeaderTableId() As String
    HeaderTableId = TextFromCodePoints("12486 12540 12502 12523") & "ID"
End Function

Private Function HeaderColumnEnglish() As String
    HeaderColumnEnglish = TextFromCodePoints("12459 12521 12512 21517 65288 33521 65289")
End Function

Private Function HeaderColumnJapanese() As String
    HeaderColumnJapanese = TextFromCodePoints("12459 12521 12512 21517 65288 21644 65289")
End Function

Private Function FieldNameHeaderText() As String
    FieldNameHeaderText = TextFromCodePoints("12501 12451 12540 12523 12489 21517")
End Function

Private Function JapaneseNameHeaderText() As String
    JapaneseNameHeaderText = TextFromCodePoints("26085 26412 35486 21517")
End Function

Private Function TextFromCodePoints(ByVal codePoints As String) As String
    Dim parts() As String
    parts = Split(codePoints, " ")

    Dim index As Long
    For index = LBound(parts) To UBound(parts)
        If Len(parts(index)) > 0 Then
            TextFromCodePoints = TextFromCodePoints & ChrW$(CLng(parts(index)))
        End If
    Next index
End Function

Private Sub FormatOutputSheet(ByVal sheet As Worksheet)
    With sheet.Rows(1)
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
    End With

    Dim lastRow As Long
    lastRow = LastUsedRowInSheet(sheet)
    If lastRow < 1 Then
        lastRow = 1
    End If

    If lastRow > FORMAT_AUTOFIT_SAMPLE_ROW_COUNT Then
        lastRow = FORMAT_AUTOFIT_SAMPLE_ROW_COUNT
    End If

    sheet.Range(sheet.Cells(1, 1), sheet.Cells(lastRow, 3)).Columns.AutoFit
    If sheet.AutoFilterMode Then
        sheet.AutoFilterMode = False
    End If
    sheet.Range("A1:C1").AutoFilter
End Sub

Private Function LastUsedRowInSheet(ByVal sheet As Worksheet) As Long
    Dim foundCell As Range
    Set foundCell = sheet.Cells.Find( _
        What:="*", _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False)

    If foundCell Is Nothing Then
        LastUsedRowInSheet = 0
    Else
        LastUsedRowInSheet = foundCell.Row
    End If
End Function

Private Function BuildCompletionMessage(ByRef stats As RunStats) As String
    BuildCompletionMessage = _
        "Analysis completed." & vbCrLf & vbCrLf & _
        "Scanned files: " & CStr(stats.ScannedFileCount) & vbCrLf & _
        "Matched files: " & CStr(stats.MatchedFileCount) & vbCrLf & _
        "Processed workbooks: " & CStr(stats.ProcessedWorkbookCount) & vbCrLf & _
        "Skipped workbooks: " & CStr(stats.SkippedWorkbookCount) & vbCrLf & _
        "Skipped folders: " & CStr(stats.SkippedFolderCount) & vbCrLf & _
        "Checkpoint saves: " & CStr(stats.CheckpointSaveCount) & vbCrLf & _
        "Checkpoint save failures: " & CStr(stats.CheckpointSaveFailureCount) & vbCrLf & _
        "Table rows: " & CStr(stats.TableCount) & vbCrLf & _
        "Column rows: " & CStr(stats.ColumnCount)
End Function
