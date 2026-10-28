Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Threading
Imports Excel = Microsoft.Office.Interop.Excel

''' <summary>
''' Handles report template operations on Excel workbooks.
''' </summary>
Public Class ExcelTemplateHandler

    ''' <summary>
    ''' Adds the cover page from a report template to an existing workbook.
    ''' </summary>
    ''' <param name="baseFile">The path to the workbook to modify.</param>
    ''' <param name="templateFile">The path to the report template.</param>
    ''' <param name="reportProgress">Callback for reporting operation progress.</param>
    ''' <param name="reportStatus">Callback for reporting operation status.</param>
    Public Sub AddCoverPage(baseFile As String, templateFile As String, reportProgress As Action(Of Integer), reportStatus As Action(Of String))
        If String.IsNullOrWhiteSpace(baseFile) OrElse Not File.Exists(baseFile) Then
            Throw New FileNotFoundException("The selected base workbook could not be found.", baseFile)
        End If

        If String.IsNullOrWhiteSpace(templateFile) OrElse Not File.Exists(templateFile) Then
            Throw New FileNotFoundException("The report template could not be found.", templateFile)
        End If

        Const coverSheetName As String = "Cover Page"

        Dim excel As Excel.Application = Nothing
        Dim targetWorkbook As Excel.Workbook = Nothing
        Dim templateWorkbook As Excel.Workbook = Nothing
        Dim coverSheet As Excel.Worksheet = Nothing
        Dim targetSheets As Excel.Sheets = Nothing
        Dim targetWorksheet As Excel.Worksheet = Nothing

        Try
            reportProgress?.Invoke(0)
            reportStatus?.Invoke("Opening selected workbook...")

            excel = New Excel.Application With {
                .Visible = False,
                .DisplayAlerts = False
            }

            targetWorkbook = excel.Workbooks.Open(baseFile)

            reportProgress?.Invoke(20)
            reportStatus?.Invoke("Checking for existing cover page...")

            Try
                For i As Integer = 1 To targetWorkbook.Worksheets.Count
                    targetWorksheet = TryCast(targetWorkbook.Worksheets(i), Excel.Worksheet)

                    If targetWorksheet IsNot Nothing AndAlso String.Equals(targetWorksheet.Name, coverSheetName, StringComparison.OrdinalIgnoreCase) Then
                        Throw New InvalidOperationException("The workbook already contains a Cover Page.")
                    End If

                    If targetWorksheet IsNot Nothing Then
                        Marshal.ReleaseComObject(targetWorksheet)
                        targetWorksheet = Nothing
                    End If
                Next
            Finally
                If targetWorksheet IsNot Nothing Then
                    Marshal.ReleaseComObject(targetWorksheet)
                    targetWorksheet = Nothing
                End If
            End Try

            reportProgress?.Invoke(35)
            reportStatus?.Invoke("Opening report template...")

            templateWorkbook = excel.Workbooks.Open(templateFile, ReadOnly:=True)
            coverSheet = TryCast(templateWorkbook.Worksheets(coverSheetName), Excel.Worksheet)

            If coverSheet Is Nothing Then
                Throw New InvalidOperationException($"The template does not contain a worksheet named '{coverSheetName}'.")
            End If

            reportProgress?.Invoke(60)
            reportStatus?.Invoke("Copying cover page...")

            targetSheets = targetWorkbook.Worksheets
            CopyHeaderWorksheet(coverSheet, targetSheets, coverSheetName)

            reportProgress?.Invoke(80)
            reportStatus?.Invoke("Saving workbook...")

            targetWorkbook.Save()

            reportProgress?.Invoke(100)
            reportStatus?.Invoke("Cover page added successfully.")
        Finally
            If targetWorksheet IsNot Nothing Then Marshal.ReleaseComObject(targetWorksheet)
            If coverSheet IsNot Nothing Then Marshal.ReleaseComObject(coverSheet)
            If targetSheets IsNot Nothing Then Marshal.ReleaseComObject(targetSheets)

            If templateWorkbook IsNot Nothing Then
                Try
                    Thread.Sleep(200)
                    templateWorkbook.Close(SaveChanges:=False)
                Catch
                    ' Continue cleanup if Excel rejects the close request.
                Finally
                    Marshal.ReleaseComObject(templateWorkbook)
                End Try
            End If

            If targetWorkbook IsNot Nothing Then
                Try
                    targetWorkbook.Close(SaveChanges:=False)
                Catch
                    ' Continue cleanup if Excel rejects the close request.
                Finally
                    Marshal.ReleaseComObject(targetWorkbook)
                End Try
            End If

            If excel IsNot Nothing Then
                Try
                    excel.Quit()
                Catch
                    ' Continue cleanup if Excel rejects the quit request.
                Finally
                    Marshal.ReleaseComObject(excel)
                End Try
            End If
        End Try
    End Sub

    ''' <summary>
    ''' Adds the summary and lookup worksheets from a report template to an existing workbook.
    ''' </summary>
    ''' <param name="baseFile">The path to the workbook to modify.</param>
    ''' <param name="templateFile">The path to the report template.</param>
    ''' <param name="reportProgress">Callback for reporting operation progress.</param>
    ''' <param name="reportStatus">Callback for reporting operation status.</param>
    Public Sub AddSummaryPage(baseFile As String, templateFile As String, reportProgress As Action(Of Integer), reportStatus As Action(Of String))
        If String.IsNullOrWhiteSpace(baseFile) OrElse Not File.Exists(baseFile) Then
            Throw New FileNotFoundException("The selected base workbook could not be found.", baseFile)
        End If

        If String.IsNullOrWhiteSpace(templateFile) OrElse Not File.Exists(templateFile) Then
            Throw New FileNotFoundException("The report template could not be found.", templateFile)
        End If

        Dim excel As Excel.Application = Nothing
        Dim targetWorkbook As Excel.Workbook = Nothing
        Dim templateWorkbook As Excel.Workbook = Nothing
        Dim targetSheets As Excel.Sheets = Nothing
        Dim templateSheets As Excel.Sheets = Nothing
        Dim targetWorksheet As Excel.Worksheet = Nothing
        Dim summarySheet As Excel.Worksheet = Nothing
        Dim lookupSheet As Excel.Worksheet = Nothing
        Dim targetSummary As Excel.Worksheet = Nothing
        Dim targetLookup As Excel.Worksheet = Nothing
        Dim summaryTable As Excel.ListObject = Nothing
        Dim testItemTable As Excel.ListObject = Nothing
        Dim slotTable As Excel.ListObject = Nothing
        Dim testItemRange As Excel.Range = Nothing
        Dim slotRange As Excel.Range = Nothing
        Dim definedName As Excel.Name = Nothing

        Try
            reportProgress?.Invoke(0)
            reportStatus?.Invoke("Opening selected workbook...")

            excel = New Excel.Application With {
                .Visible = False,
                .DisplayAlerts = False
            }

            targetWorkbook = excel.Workbooks.Open(baseFile)

            reportProgress?.Invoke(20)
            reportStatus?.Invoke("Checking for existing summary and lookup worksheets...")

            targetSheets = targetWorkbook.Worksheets

            For i As Integer = 1 To targetSheets.Count
                targetWorksheet = TryCast(targetSheets(i), Excel.Worksheet)

                If targetWorksheet IsNot Nothing AndAlso
                    (String.Equals(targetWorksheet.Name, "Summary", StringComparison.OrdinalIgnoreCase) OrElse
                     String.Equals(targetWorksheet.Name, "Lookup", StringComparison.OrdinalIgnoreCase)) Then
                    Throw New InvalidOperationException("The workbook already contains a Summary or Lookup worksheet.")
                End If

                If targetWorksheet IsNot Nothing Then
                    Marshal.ReleaseComObject(targetWorksheet)
                    targetWorksheet = Nothing
                End If
            Next

            reportProgress?.Invoke(35)
            reportStatus?.Invoke("Opening report template...")

            templateWorkbook = excel.Workbooks.Open(templateFile, ReadOnly:=True)
            templateSheets = templateWorkbook.Worksheets

            summarySheet = TryCast(templateSheets("Summary"), Excel.Worksheet)
            lookupSheet = TryCast(templateSheets("Lookup"), Excel.Worksheet)

            If summarySheet Is Nothing OrElse lookupSheet Is Nothing Then
                Throw New InvalidOperationException("The template must contain both Summary and Lookup worksheets.")
            End If

            reportProgress?.Invoke(55)
            reportStatus?.Invoke("Copying lookup worksheet...")

            CopyHeaderWorksheet(lookupSheet, targetSheets, "Lookup")

            reportProgress?.Invoke(65)
            reportStatus?.Invoke("Copying summary worksheet...")

            CopyHeaderWorksheet(summarySheet, targetSheets, "Summary")

            reportProgress?.Invoke(75)
            reportStatus?.Invoke("Repairing formulas and dropdown references...")

            targetSummary = TryCast(targetWorkbook.Worksheets("Summary"), Excel.Worksheet)
            targetLookup = TryCast(targetWorkbook.Worksheets("Lookup"), Excel.Worksheet)

            If targetSummary Is Nothing OrElse targetLookup Is Nothing Then
                Throw New InvalidOperationException("The copied Summary or Lookup worksheet could not be found.")
            End If

            summaryTable = targetSummary.ListObjects("SummaryTable")
            testItemTable = targetLookup.ListObjects("TestItemTable")
            slotTable = targetSummary.ListObjects("FlXProSlotInfoTable")

            ' Recreate the workbook-level names so they reference the destination workbook's tables.
            For Each nameToReplace As String In {"SlotLookup", "TestLookup"}
                Try
                    definedName = targetWorkbook.Names(nameToReplace)
                    definedName.Delete()
                Catch
                    ' The name may not exist in the destination workbook.
                Finally
                    If definedName IsNot Nothing Then
                        Marshal.ReleaseComObject(definedName)
                        definedName = Nothing
                    End If
                End Try
            Next

            definedName = targetWorkbook.Names.Add(
                Name:="SlotLookup",
                RefersTo:="=FlXProSlotInfoTable[Slot / Chassis]")
            Marshal.ReleaseComObject(definedName)
            definedName = Nothing

            definedName = targetWorkbook.Names.Add(
                Name:="TestLookup",
                RefersTo:="=TestItemTable[ITEMS]")
            Marshal.ReleaseComObject(definedName)
            definedName = Nothing

            ' Repair calculated columns to use local table references.
            SetTableColumnFormula(summaryTable, "DVT", "=XLOOKUP([@OMS],TestItemTable[OMS],TestItemTable[DVT],"""")")
            SetTableColumnFormula(summaryTable, "OMS", "=XLOOKUP([@[Test Item]],TestItemTable[ITEMS],TestItemTable[OMS],"""")")
            SetTableColumnFormula(summaryTable, "Qualifier", "=XLOOKUP([@[Test Item]],TestItemTable[ITEMS],TestItemTable[Qualifier],"""")")
            SetTableColumnFormula(summaryTable, "Unit", "=XLOOKUP([@[Test Item]],TestItemTable[ITEMS],TestItemTable[Unit],"""")")
            SetTableColumnFormula(summaryTable, "Model", "=XLOOKUP([@[Slot / Chassis]],FlXProSlotInfoTable[Slot / Chassis],FlXProSlotInfoTable[Model],"""")")
            SetTableColumnFormula(summaryTable, "Firmware", "=XLOOKUP([@[Slot / Chassis]],FlXProSlotInfoTable[Slot / Chassis],FlXProSlotInfoTable[Firmware],"""")")
            SetTableColumnFormula(summaryTable, "Serial Number", "=XLOOKUP([@[Slot / Chassis]],FlXProSlotInfoTable[Slot / Chassis],FlXProSlotInfoTable[Serial Number],"""")")

            ' Restore validation on the actual table columns, not the worksheet columns.
            If summaryTable.ListColumns("Test Item").DataBodyRange IsNot Nothing Then
                testItemRange = summaryTable.ListColumns("Test Item").DataBodyRange
                testItemRange.Validation.Delete()
                testItemRange.Validation.Add(
                    Type:=Microsoft.Office.Interop.Excel.XlDVType.xlValidateList,
                    AlertStyle:=Microsoft.Office.Interop.Excel.XlDVAlertStyle.xlValidAlertStop,
                    Operator:=Microsoft.Office.Interop.Excel.XlFormatConditionOperator.xlBetween,
                    Formula1:="=TestLookup")
                testItemRange.Validation.IgnoreBlank = True
                testItemRange.Validation.InCellDropdown = True
                Marshal.ReleaseComObject(testItemRange)
                testItemRange = Nothing
            End If

            If summaryTable.ListColumns("Slot / Chassis").DataBodyRange IsNot Nothing Then
                slotRange = summaryTable.ListColumns("Slot / Chassis").DataBodyRange
                slotRange.Validation.Delete()
                slotRange.Validation.Add(
                    Type:=Microsoft.Office.Interop.Excel.XlDVType.xlValidateList,
                    AlertStyle:=Microsoft.Office.Interop.Excel.XlDVAlertStyle.xlValidAlertStop,
                    Operator:=Microsoft.Office.Interop.Excel.XlFormatConditionOperator.xlBetween,
                    Formula1:="=SlotLookup")
                slotRange.Validation.IgnoreBlank = True
                slotRange.Validation.InCellDropdown = True
                Marshal.ReleaseComObject(slotRange)
                slotRange = Nothing
            End If

            reportProgress?.Invoke(90)
            reportStatus?.Invoke("Saving workbook...")

            targetWorkbook.Save()

            reportProgress?.Invoke(100)
            reportStatus?.Invoke("Summary and lookup worksheets added successfully.")
        Finally
            If testItemRange IsNot Nothing Then Marshal.ReleaseComObject(testItemRange)
            If slotRange IsNot Nothing Then Marshal.ReleaseComObject(slotRange)
            If definedName IsNot Nothing Then Marshal.ReleaseComObject(definedName)
            If summaryTable IsNot Nothing Then Marshal.ReleaseComObject(summaryTable)
            If testItemTable IsNot Nothing Then Marshal.ReleaseComObject(testItemTable)
            If slotTable IsNot Nothing Then Marshal.ReleaseComObject(slotTable)
            If targetSummary IsNot Nothing Then Marshal.ReleaseComObject(targetSummary)
            If targetLookup IsNot Nothing Then Marshal.ReleaseComObject(targetLookup)
            If targetWorksheet IsNot Nothing Then Marshal.ReleaseComObject(targetWorksheet)
            If summarySheet IsNot Nothing Then Marshal.ReleaseComObject(summarySheet)
            If lookupSheet IsNot Nothing Then Marshal.ReleaseComObject(lookupSheet)
            If templateSheets IsNot Nothing Then Marshal.ReleaseComObject(templateSheets)
            If targetSheets IsNot Nothing Then Marshal.ReleaseComObject(targetSheets)

            If templateWorkbook IsNot Nothing Then
                Try
                    Thread.Sleep(200)
                    templateWorkbook.Close(SaveChanges:=False)
                Catch
                    ' Continue cleanup if Excel rejects the close request.
                Finally
                    Marshal.ReleaseComObject(templateWorkbook)
                End Try
            End If

            If targetWorkbook IsNot Nothing Then
                Try
                    targetWorkbook.Close(SaveChanges:=False)
                Catch
                    ' Continue cleanup if Excel rejects the close request.
                Finally
                    Marshal.ReleaseComObject(targetWorkbook)
                End Try
            End If

            If excel IsNot Nothing Then
                Try
                    excel.Quit()
                Catch
                    ' Continue cleanup if Excel rejects the quit request.
                Finally
                    Marshal.ReleaseComObject(excel)
                End Try
            End If
        End Try
    End Sub

    ''' <summary>
    ''' Copies a header worksheet into the destination workbook in header-priority order.
    ''' </summary>
    Private Sub CopyHeaderWorksheet(sourceSheet As Excel.Worksheet, targetSheets As Excel.Sheets, sheetName As String)
        Dim insertionIndex As Integer = 1
        Dim insertionSheet As Excel.Worksheet = Nothing

        Try
            Dim newPriority As Integer = HeaderSequence.GetHeaderPriority(sheetName)

            For i As Integer = 1 To targetSheets.Count
                Dim worksheet As Excel.Worksheet = Nothing

                Try
                    worksheet = CType(targetSheets.Item(i), Excel.Worksheet)

                    Dim existingPriority As Integer = HeaderSequence.GetHeaderPriority(worksheet.Name)

                    ' Header sheets are expected to be grouped at the beginning.
                    If existingPriority = Integer.MaxValue Then
                        Exit For
                    End If

                    ' Insert after headers of equal or higher priority.
                    If existingPriority > newPriority Then
                        insertionIndex = i
                        Exit For
                    End If

                    insertionIndex = i + 1
                Finally
                    ExcelHelper.ReleaseComObject(worksheet)
                End Try
            Next

            If insertionIndex > targetSheets.Count Then
                insertionSheet = CType(targetSheets.Item(targetSheets.Count), Excel.Worksheet)
                sourceSheet.Copy(After:=insertionSheet)
            Else
                insertionSheet = CType(targetSheets.Item(insertionIndex), Excel.Worksheet)
                sourceSheet.Copy(Before:=insertionSheet)
            End If
        Finally
            ExcelHelper.ReleaseComObject(insertionSheet)
        End Try
    End Sub

    Private Sub SetTableColumnFormula(table As Excel.ListObject, columnName As String, formula As String)
        Dim columnRange As Excel.Range = Nothing

        Try
            columnRange = table.ListColumns(columnName).DataBodyRange

            If columnRange IsNot Nothing Then
                columnRange.Formula = formula
            End If
        Finally
            If columnRange IsNot Nothing Then Marshal.ReleaseComObject(columnRange)
        End Try
    End Sub

End Class