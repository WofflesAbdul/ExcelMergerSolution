Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel

Public NotInheritable Class ExcelWorksheetMergerReport

    Public Shared Sub MergeReportWorkbook(destinationFile As String,
                                          sourceFile As String,
                                          Optional reportProgress As Action(Of Integer) = Nothing,
                                          Optional reportStatus As Action(Of String) = Nothing)

        Using messageFilter As New ExcelInteropHelper.ExcelMessageFilterScope()
            Using session As New ExcelWorkbookSession(destinationFile, sourceFile)
                Dim sourceSheets As Excel.Sheets = Nothing
                Dim targetSheets As Excel.Sheets = Nothing
                Dim headerInserter As New ExcelWorksheetHeaderInserter()
                Dim sourceSummary As Excel.Worksheet = Nothing
                Dim targetSummary As Excel.Worksheet = Nothing
                Dim sourceLookup As Excel.Worksheet = Nothing
                Dim targetLookup As Excel.Worksheet = Nothing

                Try
                    sourceSheets = session.SourceWorkbook.Worksheets
                    targetSheets = session.DestinationWorkbook.Worksheets

                    reportProgress?.Invoke(0)
                    reportStatus?.Invoke("Merging report workbook...")

                    ' Copy the Cover Page first.
                    CopyNamedWorksheet(sourceSheets, targetSheets, "Cover Page", headerInserter, reportStatus)

                    reportProgress?.Invoke(20)

                    ' Lookup must be copied before Summary so its references can be repaired.
                    CopyNamedWorksheet(sourceSheets, targetSheets, "Lookup", headerInserter, reportStatus)

                    reportProgress?.Invoke(35)

                    ' Copy Summary if available and not already present.
                    sourceSummary = ExcelWorksheetHelper.GetWorksheetIfExists(sourceSheets, "Summary")

                    sourceLookup = ExcelWorksheetHelper.GetWorksheetIfExists(sourceSheets, "Lookup")

                    Dim summaryCopied As Boolean = False

                    If sourceSummary Is Nothing Then
                        reportStatus?.Invoke(
                            "Source workbook does not contain a Summary worksheet. Skipping copy.")
                    Else
                        summaryCopied = ExcelWorksheetHelper.CopyWorksheetIfMissing(sourceSummary, targetSheets, "Summary", headerInserter)

                        If summaryCopied Then
                            reportStatus?.Invoke("Summary copied.")
                        Else
                            reportStatus?.Invoke(
                                "Summary already exists. Skipping copy.")
                        End If
                    End If

                    reportProgress?.Invoke(50)

                    ' Repair references only when Summary was newly copied and Lookup
                    ' exists in both workbooks.
                    If summaryCopied AndAlso sourceLookup IsNot Nothing Then
                        targetLookup = ExcelWorksheetHelper.GetWorksheetIfExists(targetSheets, "Lookup")

                        If targetLookup IsNot Nothing Then
                            targetSummary = ExcelWorksheetHelper.GetWorksheetIfExists(targetSheets, "Summary")

                            If targetSummary IsNot Nothing Then
                                ExcelSummaryReferenceRepairer.RepairSummaryReferences(session.SourceWorkbook, session.DestinationWorkbook, sourceSummary, targetSummary, Path.GetFileName(sourceFile))

                                reportStatus?.Invoke("Summary references repaired.")
                            End If
                        End If
                    End If

                    reportProgress?.Invoke(60)

                    ' Copy remaining worksheets. Special worksheets are already handled;
                    ' duplicate names are skipped case-insensitively by the shared helper.
                    Dim sourceSheetCount As Integer = sourceSheets.Count

                    For index As Integer = 1 To sourceSheetCount
                        Dim sourceSheet As Excel.Worksheet = Nothing

                        Try
                            sourceSheet = CType(sourceSheets.Item(index), Excel.Worksheet)
                            Dim sheetName As String = sourceSheet.Name

                            If Not IsSpecialWorksheet(sheetName) Then
                                Dim copied As Boolean =
                                    ExcelWorksheetHelper.CopyWorksheetIfMissing(sourceSheet, targetSheets, sheetName, headerInserter)

                                If copied Then
                                    reportStatus?.Invoke(sheetName & " copied.")
                                Else
                                    reportStatus?.Invoke(sheetName & " already exists. Skipping copy.")
                                End If
                            End If
                        Finally
                            ExcelInteropHelper.ReleaseComObject(sourceSheet)
                        End Try

                        reportProgress?.Invoke(60 + CInt(index / Math.Max(1, sourceSheetCount) * 35))
                    Next

                    ' Reset worksheet cursors and save the destination once.
                    ExcelInteropHelper.ResetWorksheetCursors(session.DestinationWorkbook)
                    session.SaveDestination()

                    reportProgress?.Invoke(100)
                    reportStatus?.Invoke("Report workbook merged successfully.")

                Finally
                    ExcelInteropHelper.ReleaseComObject(targetSummary)
                    ExcelInteropHelper.ReleaseComObject(targetLookup)
                    ExcelInteropHelper.ReleaseComObject(sourceSummary)
                    ExcelInteropHelper.ReleaseComObject(sourceLookup)
                    ExcelInteropHelper.ReleaseComObject(targetSheets)
                    ExcelInteropHelper.ReleaseComObject(sourceSheets)
                End Try
            End Using
        End Using
    End Sub

    Private Shared Sub CopyNamedWorksheet(sourceSheets As Excel.Sheets,
                                          targetSheets As Excel.Sheets,
                                          sheetName As String,
                                          headerInserter As ExcelWorksheetHeaderInserter,
                                          reportStatus As Action(Of String))

        Dim sourceSheet As Excel.Worksheet = Nothing

        Try
            sourceSheet = ExcelWorksheetHelper.GetWorksheetIfExists(
                sourceSheets, sheetName)

            If sourceSheet Is Nothing Then
                reportStatus?.Invoke(
                    "Source workbook does not contain a " &
                    sheetName & " worksheet. Skipping copy.")
                Return
            End If

            Dim copied As Boolean = ExcelWorksheetHelper.CopyWorksheetIfMissing(sourceSheet, targetSheets, sheetName, headerInserter)

            If copied Then
                reportStatus?.Invoke(sheetName & " copied.")
            Else
                reportStatus?.Invoke(
                    sheetName & " already exists. Skipping copy.")
            End If
        Finally
            ExcelInteropHelper.ReleaseComObject(sourceSheet)
        End Try
    End Sub

    Private Shared Function IsSpecialWorksheet(sheetName As String) As Boolean
        Return String.Equals(sheetName, "Cover Page", StringComparison.OrdinalIgnoreCase) OrElse
               String.Equals(sheetName, "Summary", StringComparison.OrdinalIgnoreCase) OrElse
               String.Equals(sheetName, "Lookup", StringComparison.OrdinalIgnoreCase)
    End Function

End Class