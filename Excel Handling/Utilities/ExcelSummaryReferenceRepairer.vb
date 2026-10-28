Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions
Imports Excel = Microsoft.Office.Interop.Excel

''' <summary>
''' Repairs formulas, validations, and workbook-level names on a copied Summary worksheet.
''' </summary>
Public NotInheritable Class ExcelSummaryReferenceRepairer

    ''' <summary>
    ''' Repairs Summary references after copying a Summary worksheet from a source workbook.
    ''' </summary>
    Public Shared Sub RepairSummaryReferences(sourceWorkbook As Excel.Workbook, targetWorkbook As Excel.Workbook, sourceSheet As Excel.Worksheet, targetSheet As Excel.Worksheet, sourceFileName As String)
        If sourceWorkbook Is Nothing Then Throw New ArgumentNullException(NameOf(sourceWorkbook))
        If targetWorkbook Is Nothing Then Throw New ArgumentNullException(NameOf(targetWorkbook))
        If sourceSheet Is Nothing Then Throw New ArgumentNullException(NameOf(sourceSheet))
        If targetSheet Is Nothing Then Throw New ArgumentNullException(NameOf(targetSheet))

        RepairWorkbookName(sourceWorkbook, targetWorkbook, "TestLookup", sourceFileName)
        RepairWorkbookName(sourceWorkbook, targetWorkbook, "SlotLookup", sourceFileName)
        RepairSummaryFormulas(sourceSheet, targetSheet)
        RepairSummaryValidations(sourceSheet, targetSheet, sourceFileName)
    End Sub

    ''' <summary>
    ''' Copies a source workbook-level name definition into the destination workbook.
    ''' </summary>
    Private Shared Sub RepairWorkbookName(sourceWorkbook As Excel.Workbook, targetWorkbook As Excel.Workbook, nameText As String, sourceFileName As String)
        Dim sourceName As Excel.Name = Nothing
        Dim targetName As Excel.Name = Nothing
        Dim sourceNames As Excel.Names = Nothing
        Dim targetNames As Excel.Names = Nothing

        Try
            sourceNames = sourceWorkbook.Names
            targetNames = targetWorkbook.Names
            sourceName = GetWorkbookNameIfExists(sourceNames, nameText)

            If sourceName Is Nothing Then Return

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
    ''' Finds a workbook-level defined name without matching worksheet-scoped names.
    ''' </summary>
    Private Shared Function GetWorkbookNameIfExists(names As Excel.Names, nameText As String) As Excel.Name
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
    Private Shared Function NormalizeWorkbookReference(reference As String, sourceFileName As String) As String
        If String.IsNullOrEmpty(reference) Then Return reference

        Dim escapedFileName As String = Regex.Escape(sourceFileName)

        reference = Regex.Replace(
            reference,
            "'[^']*\[" & escapedFileName & "\]([^']*)'!",
            "'$1'!",
            RegexOptions.IgnoreCase)

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
    Private Shared Sub RepairSummaryFormulas(sourceSheet As Excel.Worksheet, targetSheet As Excel.Worksheet)
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
    Private Shared Sub RepairSummaryValidations(sourceSheet As Excel.Worksheet, targetSheet As Excel.Worksheet, sourceFileName As String)
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

End Class