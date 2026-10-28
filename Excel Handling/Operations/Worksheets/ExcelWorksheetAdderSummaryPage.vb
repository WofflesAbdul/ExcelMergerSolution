Imports Excel = Microsoft.Office.Interop.Excel

''' <summary>
''' Handles adding Summary and Lookup worksheets from a report template.
''' </summary>
Public Class ExcelWorksheetAdderSummaryPage

    ''' <summary>
    ''' Adds Summary and Lookup worksheets from a report template to an existing workbook.
    ''' </summary>
    Public Sub AddSummaryPage(baseFile As String, templateFile As String, reportProgress As Action(Of Integer), reportStatus As Action(Of String))
        Const summarySheetName As String = "Summary"
        Const lookupSheetName As String = "Lookup"

        Using messageFilter As New ExcelInteropHelper.ExcelMessageFilterScope()
            Using session As New ExcelWorkbookSession(baseFile, templateFile)
                Dim targetSheets As Excel.Sheets = Nothing
                Dim templateSheets As Excel.Sheets = Nothing
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
                    reportStatus?.Invoke("Opening selected workbook and report template...")

                    targetSheets = session.DestinationWorkbook.Worksheets
                    templateSheets = session.SourceWorkbook.Worksheets

                    reportProgress?.Invoke(20)
                    reportStatus?.Invoke("Checking for existing Summary and Lookup worksheets...")

                    If ExcelWorksheetHelper.WorksheetExists(targetSheets, summarySheetName) OrElse
                       ExcelWorksheetHelper.WorksheetExists(targetSheets, lookupSheetName) Then
                        Throw New InvalidOperationException("The workbook already contains a Summary or Lookup worksheet.")
                    End If

                    reportProgress?.Invoke(35)
                    reportStatus?.Invoke("Checking report template...")

                    summarySheet = ExcelWorksheetHelper.GetWorksheetIfExists(templateSheets, summarySheetName)
                    lookupSheet = ExcelWorksheetHelper.GetWorksheetIfExists(templateSheets, lookupSheetName)

                    If summarySheet Is Nothing OrElse lookupSheet Is Nothing Then
                        Throw New InvalidOperationException("The template must contain both Summary and Lookup worksheets.")
                    End If

                    Dim headerInserter As New ExcelWorksheetHeaderInserter()

                    reportProgress?.Invoke(55)
                    reportStatus?.Invoke("Copying Lookup worksheet...")

                    headerInserter.CopyHeaderWorksheet(lookupSheet, targetSheets, lookupSheetName)

                    reportProgress?.Invoke(65)
                    reportStatus?.Invoke("Copying Summary worksheet...")

                    headerInserter.CopyHeaderWorksheet(summarySheet, targetSheets, summarySheetName)

                    reportProgress?.Invoke(75)
                    reportStatus?.Invoke("Repairing formulas and dropdown references...")

                    targetSummary = ExcelWorksheetHelper.GetWorksheetIfExists(targetSheets, summarySheetName)
                    targetLookup = ExcelWorksheetHelper.GetWorksheetIfExists(targetSheets, lookupSheetName)

                    If targetSummary Is Nothing OrElse targetLookup Is Nothing Then
                        Throw New InvalidOperationException("The copied Summary or Lookup worksheet could not be found.")
                    End If

                    summaryTable = targetSummary.ListObjects("SummaryTable")
                    testItemTable = targetLookup.ListObjects("TestItemTable")
                    slotTable = targetSummary.ListObjects("FlXProSlotInfoTable")

                    For Each nameToReplace As String In {"SlotLookup", "TestLookup"}
                        Try
                            definedName = session.DestinationWorkbook.Names(nameToReplace)
                            definedName.Delete()
                        Catch
                            ' The workbook-level name may not exist.
                        Finally
                            ExcelInteropHelper.ReleaseComObject(definedName)
                            definedName = Nothing
                        End Try
                    Next

                    definedName = session.DestinationWorkbook.Names.Add(
                        Name:="SlotLookup",
                        RefersTo:="=FlXProSlotInfoTable[Slot / Chassis]")
                    ExcelInteropHelper.ReleaseComObject(definedName)
                    definedName = Nothing

                    definedName = session.DestinationWorkbook.Names.Add(
                        Name:="TestLookup",
                        RefersTo:="=TestItemTable[ITEMS]")
                    ExcelInteropHelper.ReleaseComObject(definedName)
                    definedName = Nothing

                    SetTableColumnFormula(summaryTable, "DVT", "=XLOOKUP([@OMS],TestItemTable[OMS],TestItemTable[DVT],"""")")
                    SetTableColumnFormula(summaryTable, "OMS", "=XLOOKUP([@[Test Item]],TestItemTable[ITEMS],TestItemTable[OMS],"""")")
                    SetTableColumnFormula(summaryTable, "Qualifier", "=XLOOKUP([@[Test Item]],TestItemTable[ITEMS],TestItemTable[Qualifier],"""")")
                    SetTableColumnFormula(summaryTable, "Unit", "=XLOOKUP([@[Test Item]],TestItemTable[ITEMS],TestItemTable[Unit],"""")")
                    SetTableColumnFormula(summaryTable, "Model", "=XLOOKUP([@[Slot / Chassis]],FlXProSlotInfoTable[Slot / Chassis],FlXProSlotInfoTable[Model],"""")")
                    SetTableColumnFormula(summaryTable, "Firmware", "=XLOOKUP([@[Slot / Chassis]],FlXProSlotInfoTable[Slot / Chassis],FlXProSlotInfoTable[Firmware],"""")")
                    SetTableColumnFormula(summaryTable, "Serial Number", "=XLOOKUP([@[Slot / Chassis]],FlXProSlotInfoTable[Slot / Chassis],FlXProSlotInfoTable[Serial Number],"""")")

                    If summaryTable.ListColumns("Test Item").DataBodyRange IsNot Nothing Then
                        testItemRange = summaryTable.ListColumns("Test Item").DataBodyRange
                        testItemRange.Validation.Delete()
                        testItemRange.Validation.Add(
                            Type:=Excel.XlDVType.xlValidateList,
                            AlertStyle:=Excel.XlDVAlertStyle.xlValidAlertStop,
                            Operator:=Excel.XlFormatConditionOperator.xlBetween,
                            Formula1:="=TestLookup")
                        testItemRange.Validation.IgnoreBlank = True
                        testItemRange.Validation.InCellDropdown = True
                        ExcelInteropHelper.ReleaseComObject(testItemRange)
                        testItemRange = Nothing
                    End If

                    If summaryTable.ListColumns("Slot / Chassis").DataBodyRange IsNot Nothing Then
                        slotRange = summaryTable.ListColumns("Slot / Chassis").DataBodyRange
                        slotRange.Validation.Delete()
                        slotRange.Validation.Add(
                            Type:=Excel.XlDVType.xlValidateList,
                            AlertStyle:=Excel.XlDVAlertStyle.xlValidAlertStop,
                            Operator:=Excel.XlFormatConditionOperator.xlBetween,
                            Formula1:="=SlotLookup")
                        slotRange.Validation.IgnoreBlank = True
                        slotRange.Validation.InCellDropdown = True
                        ExcelInteropHelper.ReleaseComObject(slotRange)
                        slotRange = Nothing
                    End If

                    reportProgress?.Invoke(90)
                    reportStatus?.Invoke("Saving workbook...")

                    session.SaveDestination()

                    reportProgress?.Invoke(100)
                    reportStatus?.Invoke("Summary and Lookup worksheets added successfully.")
                Finally
                    ExcelInteropHelper.ReleaseComObject(testItemRange)
                    ExcelInteropHelper.ReleaseComObject(slotRange)
                    ExcelInteropHelper.ReleaseComObject(definedName)
                    ExcelInteropHelper.ReleaseComObject(summaryTable)
                    ExcelInteropHelper.ReleaseComObject(testItemTable)
                    ExcelInteropHelper.ReleaseComObject(slotTable)
                    ExcelInteropHelper.ReleaseComObject(targetSummary)
                    ExcelInteropHelper.ReleaseComObject(targetLookup)
                    ExcelInteropHelper.ReleaseComObject(summarySheet)
                    ExcelInteropHelper.ReleaseComObject(lookupSheet)
                    ExcelInteropHelper.ReleaseComObject(templateSheets)
                    ExcelInteropHelper.ReleaseComObject(targetSheets)
                End Try
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' Sets a table column formula when the column contains data rows.
    ''' </summary>
    Private Sub SetTableColumnFormula(table As Excel.ListObject, columnName As String, formula As String)
        Dim columnRange As Excel.Range = Nothing

        Try
            columnRange = table.ListColumns(columnName).DataBodyRange

            If columnRange IsNot Nothing Then
                columnRange.Formula = formula
            End If
        Finally
            ExcelInteropHelper.ReleaseComObject(columnRange)
        End Try
    End Sub

End Class