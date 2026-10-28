Imports Excel = Microsoft.Office.Interop.Excel

''' <summary>
''' Handles adding a Cover Page worksheet from a report template.
''' </summary>
Public Class ExcelWorksheetAdderCoverPage

    ''' <summary>
    ''' Adds the Cover Page from a report template to an existing workbook.
    ''' </summary>
    Public Sub AddCoverPage(baseFile As String, templateFile As String, reportProgress As Action(Of Integer), reportStatus As Action(Of String))
        Const sheetName As String = "Cover Page"

        Using messageFilter As New ExcelInteropHelper.ExcelMessageFilterScope()
            Using session As New ExcelWorkbookSession(baseFile, templateFile)
                Dim targetSheets As Excel.Sheets = Nothing
                Dim templateSheets As Excel.Sheets = Nothing
                Dim sourceSheet As Excel.Worksheet = Nothing

                Try
                    reportProgress?.Invoke(0)
                    reportStatus?.Invoke("Opening selected workbook and report template...")

                    targetSheets = session.DestinationWorkbook.Worksheets
                    templateSheets = session.SourceWorkbook.Worksheets

                    reportProgress?.Invoke(20)
                    reportStatus?.Invoke("Checking for an existing Cover Page...")

                    If ExcelWorksheetHelper.WorksheetExists(targetSheets, sheetName) Then
                        Throw New InvalidOperationException("The workbook already contains a Cover Page.")
                    End If

                    reportProgress?.Invoke(35)
                    reportStatus?.Invoke("Checking report template...")

                    sourceSheet = ExcelWorksheetHelper.GetWorksheetIfExists(templateSheets, sheetName)

                    If sourceSheet Is Nothing Then
                        Throw New InvalidOperationException("The template does not contain a worksheet named '" & sheetName & "'.")
                    End If

                    reportProgress?.Invoke(60)
                    reportStatus?.Invoke("Copying Cover Page...")

                    Dim headerInserter As New ExcelWorksheetHeaderInserter()
                    headerInserter.CopyHeaderWorksheet(sourceSheet, targetSheets, sheetName)

                    reportProgress?.Invoke(80)
                    reportStatus?.Invoke("Saving workbook...")

                    session.SaveDestination()

                    reportProgress?.Invoke(100)
                    reportStatus?.Invoke("Cover Page added successfully.")
                Finally
                    ExcelInteropHelper.ReleaseComObject(sourceSheet)
                    ExcelInteropHelper.ReleaseComObject(templateSheets)
                    ExcelInteropHelper.ReleaseComObject(targetSheets)
                End Try
            End Using
        End Using
    End Sub

End Class