Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel

''' <summary>
''' Handles merging Summary and Lookup worksheets from a source workbook.
''' </summary>
Public Class ExcelWorksheetMergerSummaryPage

    ''' <summary>
    ''' Copies Summary and Lookup worksheets when needed and repairs references on a newly copied Summary.
    ''' </summary>
    Public Sub MergeSummaryPage(baseFile As String, sourceFile As String, reportProgress As Action(Of Integer), reportStatus As Action(Of String))
        Const summarySheetName As String = "Summary"
        Const lookupSheetName As String = "Lookup"

        Using messageFilter As New ExcelInteropHelper.ExcelMessageFilterScope()
            Using session As New ExcelWorkbookSession(baseFile, sourceFile)
                Dim targetSheets As Excel.Sheets = Nothing
                Dim sourceSheets As Excel.Sheets = Nothing
                Dim sourceSummary As Excel.Worksheet = Nothing
                Dim sourceLookup As Excel.Worksheet = Nothing
                Dim targetSummary As Excel.Worksheet = Nothing
                Dim targetLookup As Excel.Worksheet = Nothing

                Try
                    reportProgress?.Invoke(0)
                    reportStatus?.Invoke("Opening destination and source workbooks...")

                    targetSheets = session.DestinationWorkbook.Worksheets
                    sourceSheets = session.SourceWorkbook.Worksheets

                    reportProgress?.Invoke(20)
                    reportStatus?.Invoke("Checking source worksheets...")

                    sourceSummary = ExcelWorksheetHelper.GetWorksheetIfExists(sourceSheets, summarySheetName)
                    sourceLookup = ExcelWorksheetHelper.GetWorksheetIfExists(sourceSheets, lookupSheetName)

                    Dim headerInserter As New ExcelWorksheetHeaderInserter()
                    Dim lookupCopied As Boolean = False
                    Dim summaryCopied As Boolean = False

                    reportProgress?.Invoke(35)

                    If sourceLookup Is Nothing Then
                        reportStatus?.Invoke("Source workbook does not contain a Lookup worksheet. Skipping copy.")
                    ElseIf ExcelWorksheetHelper.WorksheetExists(targetSheets, lookupSheetName) Then
                        reportStatus?.Invoke("Lookup worksheet already exists. Skipping copy.")
                    Else
                        reportStatus?.Invoke("Copying Lookup worksheet...")
                        lookupCopied = ExcelWorksheetHelper.CopyWorksheetIfMissing(sourceLookup, targetSheets, lookupSheetName, headerInserter)
                    End If

                    reportProgress?.Invoke(55)

                    If sourceSummary Is Nothing Then
                        reportStatus?.Invoke("Source workbook does not contain a Summary worksheet. Skipping copy.")
                    ElseIf ExcelWorksheetHelper.WorksheetExists(targetSheets, summarySheetName) Then
                        reportStatus?.Invoke("Summary worksheet already exists. Skipping copy.")
                    Else
                        reportStatus?.Invoke("Copying Summary worksheet...")
                        summaryCopied = ExcelWorksheetHelper.CopyWorksheetIfMissing(sourceSummary, targetSheets, summarySheetName, headerInserter)
                    End If

                    reportProgress?.Invoke(75)

                    If summaryCopied AndAlso sourceLookup IsNot Nothing AndAlso
                       ExcelWorksheetHelper.WorksheetExists(targetSheets, lookupSheetName) Then

                        targetSummary = ExcelWorksheetHelper.GetWorksheetIfExists(targetSheets, summarySheetName)
                        targetLookup = ExcelWorksheetHelper.GetWorksheetIfExists(targetSheets, lookupSheetName)

                        If targetSummary Is Nothing OrElse targetLookup Is Nothing Then
                            Throw New InvalidOperationException("The copied Summary or Lookup worksheet could not be found.")
                        End If

                        reportStatus?.Invoke("Repairing Summary formulas and dropdown references...")

                        ExcelSummaryReferenceRepairer.RepairSummaryReferences(
                            session.SourceWorkbook,
                            session.DestinationWorkbook,
                            sourceSummary,
                            targetSummary,
                            Path.GetFileName(sourceFile))
                    End If

                    reportProgress?.Invoke(90)
                    reportStatus?.Invoke("Saving destination workbook...")

                    session.SaveDestination()

                    reportProgress?.Invoke(100)

                    If summaryCopied OrElse lookupCopied Then
                        reportStatus?.Invoke("Summary and Lookup merge completed successfully.")
                    Else
                        reportStatus?.Invoke("No worksheets needed to be copied.")
                    End If
                Finally
                    ExcelInteropHelper.ReleaseComObject(targetLookup)
                    ExcelInteropHelper.ReleaseComObject(targetSummary)
                    ExcelInteropHelper.ReleaseComObject(sourceLookup)
                    ExcelInteropHelper.ReleaseComObject(sourceSummary)
                    ExcelInteropHelper.ReleaseComObject(sourceSheets)
                    ExcelInteropHelper.ReleaseComObject(targetSheets)
                End Try
            End Using
        End Using
    End Sub

End Class