Imports Excel = Microsoft.Office.Interop.Excel

''' <summary>
''' Handles merging a Cover Page worksheet from a source workbook.
''' </summary>
Public Class ExcelWorksheetMergerCoverPage

    ''' <summary>
    ''' Copies the Cover Page into the destination workbook if it does not already exist.
    ''' </summary>
    Public Sub MergeCoverPage(baseFile As String, sourceFile As String, reportProgress As Action(Of Integer), reportStatus As Action(Of String))
        Const sheetName As String = "Cover Page"

        Using messageFilter As New ExcelInteropHelper.ExcelMessageFilterScope()
            Using session As New ExcelWorkbookSession(baseFile, sourceFile)
                Dim targetSheets As Excel.Sheets = Nothing
                Dim sourceSheets As Excel.Sheets = Nothing
                Dim sourceSheet As Excel.Worksheet = Nothing

                Try
                    reportProgress?.Invoke(0)
                    reportStatus?.Invoke("Opening destination and source workbooks...")

                    targetSheets = session.DestinationWorkbook.Worksheets

                    reportProgress?.Invoke(20)
                    reportStatus?.Invoke("Checking for an existing Cover Page...")

                    If ExcelWorksheetHelper.WorksheetExists(targetSheets, sheetName) Then
                        reportStatus?.Invoke("Cover Page already exists. Skipping copy.")
                        reportProgress?.Invoke(100)
                        Return
                    End If

                    reportProgress?.Invoke(40)
                    reportStatus?.Invoke("Checking source workbook...")

                    sourceSheets = session.SourceWorkbook.Worksheets
                    sourceSheet = ExcelWorksheetHelper.GetWorksheetIfExists(sourceSheets, sheetName)

                    If sourceSheet Is Nothing Then
                        reportStatus?.Invoke("Source workbook does not contain a Cover Page. Skipping copy.")
                        reportProgress?.Invoke(100)
                        Return
                    End If

                    reportProgress?.Invoke(60)
                    reportStatus?.Invoke("Copying Cover Page...")

                    Dim headerInserter As New ExcelWorksheetHeaderInserter()
                    headerInserter.CopyHeaderWorksheet(sourceSheet, targetSheets, sheetName)

                    reportProgress?.Invoke(80)
                    reportStatus?.Invoke("Saving destination workbook...")

                    session.SaveDestination()

                    reportProgress?.Invoke(100)
                    reportStatus?.Invoke("Cover Page merged successfully.")
                Finally
                    ExcelInteropHelper.ReleaseComObject(sourceSheet)
                    ExcelInteropHelper.ReleaseComObject(sourceSheets)
                    ExcelInteropHelper.ReleaseComObject(targetSheets)
                End Try
            End Using
        End Using
    End Sub

End Class