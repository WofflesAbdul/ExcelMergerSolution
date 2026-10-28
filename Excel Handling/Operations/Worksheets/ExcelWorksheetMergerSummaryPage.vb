Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions
Imports Excel = Microsoft.Office.Interop.Excel

Public Class ExcelWorksheetMergerSummaryPage

    Public Sub MergeSummaryPage(baseFile As String, sourceFile As String, reportProgress As Action(Of Integer), reportStatus As Action(Of String))
        Const summarySheetName As String = "Summary"
        Const lookupSheetName As String = "Lookup"

        Dim excel As Excel.Application = Nothing
        Dim targetWorkbook As Excel.Workbook = Nothing
        Dim sourceWorkbook As Excel.Workbook = Nothing
        Dim targetSheets As Excel.Sheets = Nothing
        Dim sourceSheets As Excel.Sheets = Nothing
        Dim sourceSummary As Excel.Worksheet = Nothing
        Dim sourceLookup As Excel.Worksheet = Nothing
        Dim targetSummary As Excel.Worksheet = Nothing
        Dim targetLookup As Excel.Worksheet = Nothing
        Dim summaryCopied As Boolean = False
        Dim sourceHasLookup As Boolean = False
        Dim headerInserter As New ExcelWorksheetHeaderInserter()

        If String.IsNullOrWhiteSpace(baseFile) OrElse Not File.Exists(baseFile) Then
            Throw New FileNotFoundException("The destination workbook could not be found.", baseFile)
        End If

        If String.IsNullOrWhiteSpace(sourceFile) OrElse Not File.Exists(sourceFile) Then
            Throw New FileNotFoundException("The source workbook could not be found.", sourceFile)
        End If

        If String.Equals(Path.GetFullPath(baseFile), Path.GetFullPath(sourceFile), StringComparison.OrdinalIgnoreCase) Then
            Throw New ArgumentException("The source and destination workbooks must be different files.")
        End If

        Using messageFilter As New ExcelInteropHelper.ExcelMessageFilterScope()
            Try
                reportProgress?.Invoke(0)
                reportStatus?.Invoke("Opening destination workbook...")

                excel = New Excel.Application With {
                    .Visible = False,
                    .DisplayAlerts = False
                }

                targetWorkbook = excel.Workbooks.Open(Path.GetFullPath(baseFile))
                targetSheets = targetWorkbook.Worksheets

                reportProgress?.Invoke(15)
                reportStatus?.Invoke("Opening source workbook...")

                sourceWorkbook = excel.Workbooks.Open(Path.GetFullPath(sourceFile), ReadOnly:=True)
                sourceSheets = sourceWorkbook.Worksheets

                sourceSummary = GetWorksheetIfExists(sourceSheets, summarySheetName)
                sourceLookup = GetWorksheetIfExists(sourceSheets, lookupSheetName)
                sourceHasLookup = sourceLookup IsNot Nothing

                If sourceSummary Is Nothing Then
                    Throw New InvalidOperationException("The source workbook does not contain a Summary worksheet.")
                End If

                reportProgress?.Invoke(30)
                reportStatus?.Invoke("Checking existing worksheets...")

                Dim targetSummaryExists As Boolean = WorksheetExists(targetSheets, summarySheetName)
                Dim targetLookupExists As Boolean = WorksheetExists(targetSheets, lookupSheetName)

                If Not targetLookupExists AndAlso sourceHasLookup Then
                    headerInserter.CopyHeaderWorksheet(sourceLookup, targetSheets, lookupSheetName)
                    targetLookupExists = True
                End If

                reportProgress?.Invoke(50)

                If Not targetSummaryExists Then
                    headerInserter.CopyHeaderWorksheet(sourceSummary, targetSheets, summarySheetName)
                    summaryCopied = True
                Else
                    reportStatus?.Invoke("Skipping existing Summary worksheet...")
                End If

                reportProgress?.Invoke(70)

                If summaryCopied AndAlso sourceHasLookup AndAlso targetLookupExists Then
                    targetSummary = GetWorksheetIfExists(targetSheets, summarySheetName)
                    targetLookup = GetWorksheetIfExists(targetSheets, lookupSheetName)

                    If targetSummary Is Nothing OrElse targetLookup Is Nothing Then
                        Throw New InvalidOperationException("The copied Summary or Lookup worksheet could not be found.")
                    End If

                    RepairSummaryReferences(sourceWorkbook, targetWorkbook, sourceSummary, targetSummary, Path.GetFileName(sourceFile))
                End If

                reportProgress?.Invoke(90)
                reportStatus?.Invoke("Saving destination workbook...")

                targetWorkbook.Save()

                reportProgress?.Invoke(100)
                reportStatus?.Invoke("Summary worksheet merge completed successfully.")
            Finally
                ExcelInteropHelper.ReleaseComObject(targetLookup)
                ExcelInteropHelper.ReleaseComObject(targetSummary)
                ExcelInteropHelper.ReleaseComObject(sourceLookup)
                ExcelInteropHelper.ReleaseComObject(sourceSummary)
                ExcelInteropHelper.ReleaseComObject(sourceSheets)
                ExcelInteropHelper.ReleaseComObject(targetSheets)

                If sourceWorkbook IsNot Nothing Then
                    Try
                        sourceWorkbook.Close(SaveChanges:=False)
                    Catch
                    End Try
                End If

                ExcelInteropHelper.ReleaseComObject(sourceWorkbook)

                If targetWorkbook IsNot Nothing Then
                    Try
                        targetWorkbook.Close(SaveChanges:=False)
                    Catch
                    End Try
                End If

                ExcelInteropHelper.ReleaseComObject(targetWorkbook)

                If excel IsNot Nothing Then
                    Try
                        excel.Quit()
                    Catch
                    End Try
                End If

                ExcelInteropHelper.ReleaseComObject(excel)
            End Try
        End Using
    End Sub

    ''' <summary>
    ''' Repairs formulas, workbook-level lookup names, and validation references on a newly copied Summary worksheet.
    ''' </summary>
    Private Sub RepairSummaryReferences(sourceWorkbook As Excel.Workbook, targetWorkbook As Excel.Workbook, sourceSheet As Excel.Worksheet, targetSheet As Excel.Worksheet, sourceFileName As String)
        RepairWorkbookName(sourceWorkbook, targetWorkbook, "TestLookup", sourceFileName)
        RepairWorkbookName(sourceWorkbook, targetWorkbook, "SlotLookup", sourceFileName)
        RepairSummaryFormulas(sourceSheet, targetSheet)
        RepairSummaryValidations(sourceSheet, targetSheet, sourceFileName)
    End Sub

    ''' <summary>
    ''' Copies a source workbook-level name definition into the destination workbook.
    ''' </summary>
    Private Sub RepairWorkbookName(sourceWorkbook As Excel.Workbook, targetWorkbook As Excel.Workbook, nameText As String, sourceFileName As String)
        Dim sourceName As Excel.Name = Nothing
        Dim targetName As Excel.Name = Nothing
        Dim sourceNames As Excel.Names = Nothing
        Dim targetNames As Excel.Names = Nothing

        Try
            sourceNames = sourceWorkbook.Names
            targetNames = targetWorkbook.Names
            sourceName = GetWorkbookNameIfExists(sourceNames, nameText)

            If sourceName Is Nothing Then
                Return
            End If

            Dim refersTo As String = NormalizeWorkbookReference(Convert.ToString(sourceName.RefersTo), sourceFileName)
            targetName = GetWorkbookNameIfExists(targetNames, nameText)

            If targetName Is Nothing Then
                targetName = targetWorkbook.Names.Add(Name:=nameText, RefersTo:=refersTo)
            ElseIf Not String.Equals(Convert.ToString(targetName.RefersTo), refersTo, StringComparison.Ordinal) Then
                targetName.RefersTo = refersTo
            End If
        Finally
            ExcelInteropHelper.ReleaseComObject(targetName)
            ExcelInteropHelper.ReleaseComObject(sourceName)
            ExcelInteropHelper.ReleaseComObject(targetNames)
            ExcelInteropHelper.ReleaseComObject(sourceNames)
        End Try
    End Sub

    ''' <summary>
    ''' Finds an exact workbook-level defined name without matching worksheet-scoped names.
    ''' </summary>
    Private Function GetWorkbookNameIfExists(names As Excel.Names, nameText As String) As Excel.Name
        For i As Integer = 1 To names.Count
            Dim definedName As Excel.Name = Nothing

            Try
                definedName = CType(names.Item(i), Excel.Name)

                If String.Equals(Convert.ToString(definedName.Name), nameText, StringComparison.OrdinalIgnoreCase) Then
                    Dim result As Excel.Name = definedName
                    definedName = Nothing
                    Return result
                End If
            Finally
                ExcelInteropHelper.ReleaseComObject(definedName)
            End Try
        Next

        Return Nothing
    End Function

    ''' <summary>
    ''' Removes source workbook qualifiers while retaining the referenced worksheet or range.
    ''' </summary>
    Private Function NormalizeWorkbookReference(reference As String, sourceFileName As String) As String
        If String.IsNullOrEmpty(reference) Then
            Return reference
        End If

        Dim escapedFileName As String = Regex.Escape(sourceFileName)

        ' Handles references such as 'C:\Folder\[Source.xlsx]Lookup'!$A$1:$A$10.
        reference = Regex.Replace(
            reference,
            "'[^']*\[" & escapedFileName & "\]([^']*)'!",
            "'$1'!",
            RegexOptions.IgnoreCase)

        ' Handles remaining workbook qualifiers such as [Source.xlsx]Lookup!$A$1.
        reference = Regex.Replace(
            reference,
            "\[" & escapedFileName & "\]",
            String.Empty,
            RegexOptions.IgnoreCase)

        Return reference
    End Function

    ''' <summary>
    ''' Restores formulas changed by Excel during worksheet copying.
    ''' </summary>
    Private Sub RepairSummaryFormulas(sourceSheet As Excel.Worksheet, targetSheet As Excel.Worksheet)
        Dim sourceUsedRange As Excel.Range = Nothing
        Dim sourceFormulaCells As Excel.Range = Nothing

        Try
            sourceUsedRange = sourceSheet.UsedRange

            Try
                sourceFormulaCells = sourceUsedRange.SpecialCells(Excel.XlCellType.xlCellTypeFormulas)
            Catch ex As COMException
                Return
            End Try

            For Each sourceCell As Excel.Range In sourceFormulaCells.Cells
                Dim targetCell As Excel.Range = Nothing

                Try
                    targetCell = targetSheet.Range(sourceCell.Address)

                    Dim sourceFormula As String = Convert.ToString(sourceCell.Formula)
                    Dim targetFormula As String = Convert.ToString(targetCell.Formula)

                    If Not String.Equals(sourceFormula, targetFormula, StringComparison.Ordinal) Then
                        targetCell.Formula = sourceFormula
                    End If
                Finally
                    ExcelInteropHelper.ReleaseComObject(targetCell)
                    ExcelInteropHelper.ReleaseComObject(sourceCell)
                End Try
            Next
        Finally
            ExcelInteropHelper.ReleaseComObject(sourceFormulaCells)
            ExcelInteropHelper.ReleaseComObject(sourceUsedRange)
        End Try
    End Sub

    ''' <summary>
    ''' Restores validation sources from the source worksheet while retaining validation settings.
    ''' </summary>
    Private Sub RepairSummaryValidations(sourceSheet As Excel.Worksheet, targetSheet As Excel.Worksheet, sourceFileName As String)
        Dim sourceUsedRange As Excel.Range = Nothing
        Dim sourceValidationCells As Excel.Range = Nothing

        Try
            sourceUsedRange = sourceSheet.UsedRange

            Try
                sourceValidationCells = sourceUsedRange.SpecialCells(Excel.XlCellType.xlCellTypeAllValidation)
            Catch ex As COMException
                Return
            End Try

            For Each sourceCell As Excel.Range In sourceValidationCells.Cells
                Dim targetCell As Excel.Range = Nothing
                Dim sourceValidation As Excel.Validation = Nothing
                Dim targetValidation As Excel.Validation = Nothing

                Try
                    targetCell = targetSheet.Range(sourceCell.Address)
                    sourceValidation = sourceCell.Validation
                    targetValidation = targetCell.Validation

                    Dim sourceFormula1 As String = NormalizeWorkbookReference(Convert.ToString(sourceValidation.Formula1), sourceFileName)
                    Dim sourceFormula2 As String = NormalizeWorkbookReference(Convert.ToString(sourceValidation.Formula2), sourceFileName)
                    Dim targetFormula1 As String = Convert.ToString(targetValidation.Formula1)
                    Dim targetFormula2 As String = Convert.ToString(targetValidation.Formula2)

                    If Not String.Equals(sourceFormula1, targetFormula1, StringComparison.Ordinal) OrElse
                       Not String.Equals(sourceFormula2, targetFormula2, StringComparison.Ordinal) Then

                        targetValidation.Modify(
                            Type:=sourceValidation.Type,
                            AlertStyle:=sourceValidation.AlertStyle,
                            Operator:=sourceValidation.Operator,
                            Formula1:=sourceFormula1,
                            Formula2:=sourceFormula2)

                        targetValidation.IgnoreBlank = sourceValidation.IgnoreBlank
                        targetValidation.InCellDropdown = sourceValidation.InCellDropdown
                        targetValidation.ShowInput = sourceValidation.ShowInput
                        targetValidation.ShowError = sourceValidation.ShowError

                        targetValidation.InputTitle = sourceValidation.InputTitle
                        targetValidation.InputMessage = sourceValidation.InputMessage
                        targetValidation.ErrorTitle = sourceValidation.ErrorTitle
                        targetValidation.ErrorMessage = sourceValidation.ErrorMessage
                    End If
                Finally
                    ExcelInteropHelper.ReleaseComObject(targetValidation)
                    ExcelInteropHelper.ReleaseComObject(sourceValidation)
                    ExcelInteropHelper.ReleaseComObject(targetCell)
                    ExcelInteropHelper.ReleaseComObject(sourceCell)
                End Try
            Next
        Finally
            ExcelInteropHelper.ReleaseComObject(sourceValidationCells)
            ExcelInteropHelper.ReleaseComObject(sourceUsedRange)
        End Try
    End Sub

    Private Function GetWorksheetIfExists(sheets As Excel.Sheets, sheetName As String) As Excel.Worksheet
        For i As Integer = 1 To sheets.Count
            Dim worksheet As Excel.Worksheet = Nothing

            Try
                worksheet = CType(sheets.Item(i), Excel.Worksheet)

                If String.Equals(worksheet.Name, sheetName, StringComparison.OrdinalIgnoreCase) Then
                    Dim result As Excel.Worksheet = worksheet
                    worksheet = Nothing
                    Return result
                End If
            Finally
                ExcelInteropHelper.ReleaseComObject(worksheet)
            End Try
        Next

        Return Nothing
    End Function

    Private Function WorksheetExists(sheets As Excel.Sheets, sheetName As String) As Boolean
        Dim worksheet As Excel.Worksheet = Nothing

        Try
            worksheet = GetWorksheetIfExists(sheets, sheetName)
            Return worksheet IsNot Nothing
        Finally
            ExcelInteropHelper.ReleaseComObject(worksheet)
        End Try
    End Function

End Class