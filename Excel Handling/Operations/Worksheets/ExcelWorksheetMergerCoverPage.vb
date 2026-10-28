Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel

''' <summary>
''' Merges a Cover Page worksheet from a source workbook into a destination workbook.
''' </summary>
Public Class ExcelWorksheetMergerCoverPage

    ''' <summary>
    ''' Copies the Cover Page from the source workbook into the destination workbook.
    ''' </summary>
    ''' <param name="baseFile">The destination workbook path.</param>
    ''' <param name="sourceFile">The source workbook path.</param>
    ''' <param name="reportProgress">Callback for reporting operation progress.</param>
    ''' <param name="reportStatus">Callback for reporting operation status.</param>
    Public Sub MergeCoverPage(baseFile As String, sourceFile As String, reportProgress As Action(Of Integer), reportStatus As Action(Of String))
        If String.IsNullOrWhiteSpace(baseFile) OrElse Not File.Exists(baseFile) Then
            Throw New FileNotFoundException("The destination workbook could not be found.", baseFile)
        End If

        If String.IsNullOrWhiteSpace(sourceFile) OrElse Not File.Exists(sourceFile) Then
            Throw New FileNotFoundException("The source workbook could not be found.", sourceFile)
        End If

        If String.Equals(Path.GetFullPath(baseFile), Path.GetFullPath(sourceFile), StringComparison.OrdinalIgnoreCase) Then
            Throw New InvalidOperationException("The source and destination workbooks must be different files.")
        End If

        Const coverSheetName As String = "Cover Page"

        Dim excel As Excel.Application = Nothing
        Dim targetWorkbook As Excel.Workbook = Nothing
        Dim sourceWorkbook As Excel.Workbook = Nothing
        Dim targetSheets As Excel.Sheets = Nothing
        Dim targetWorksheet As Excel.Worksheet = Nothing
        Dim sourceSheet As Excel.Worksheet = Nothing

        Using messageFilter As New ExcelInteropHelper.ExcelMessageFilterScope()
            Try
                reportProgress?.Invoke(0)
                reportStatus?.Invoke("Opening destination workbook...")

                excel = New Excel.Application With {
                    .Visible = False,
                    .DisplayAlerts = False
                }

                targetWorkbook = excel.Workbooks.Open(baseFile)

                reportProgress?.Invoke(20)
                reportStatus?.Invoke("Checking for an existing Cover Page...")

                targetSheets = targetWorkbook.Worksheets

                For i As Integer = 1 To targetSheets.Count
                    targetWorksheet = TryCast(targetSheets.Item(i), Excel.Worksheet)

                    If targetWorksheet IsNot Nothing Then
                        Dim existingName As String = targetWorksheet.Name
                        ExcelInteropHelper.ReleaseComObject(targetWorksheet)
                        targetWorksheet = Nothing

                        If String.Equals(existingName, coverSheetName, StringComparison.OrdinalIgnoreCase) Then
                            Throw New InvalidOperationException("The destination workbook already contains a Cover Page.")
                        End If
                    End If
                Next

                reportProgress?.Invoke(40)
                reportStatus?.Invoke("Opening source workbook...")

                sourceWorkbook = excel.Workbooks.Open(sourceFile, ReadOnly:=True)

                Try
                    sourceSheet = TryCast(sourceWorkbook.Worksheets(coverSheetName), Excel.Worksheet)
                Catch
                    Throw New InvalidOperationException("The source workbook does not contain a worksheet named 'Cover Page'.")
                End Try

                If sourceSheet Is Nothing Then
                    Throw New InvalidOperationException("The source workbook does not contain a worksheet named 'Cover Page'.")
                End If

                reportProgress?.Invoke(60)
                reportStatus?.Invoke("Copying Cover Page...")

                Dim headerInserter As New ExcelWorksheetHeaderInserter()
                headerInserter.CopyHeaderWorksheet(sourceSheet, targetSheets, coverSheetName)

                reportProgress?.Invoke(80)
                reportStatus?.Invoke("Saving destination workbook...")

                targetWorkbook.Save()

                reportProgress?.Invoke(100)
                reportStatus?.Invoke("Cover Page merged successfully.")
            Finally
                ExcelInteropHelper.ReleaseComObject(targetWorksheet)
                ExcelInteropHelper.ReleaseComObject(sourceSheet)
                ExcelInteropHelper.ReleaseComObject(targetSheets)

                If sourceWorkbook IsNot Nothing Then
                    Try
                        sourceWorkbook.Close(SaveChanges:=False)
                    Catch
                        ' Continue cleanup if Excel rejects the close request.
                    Finally
                        ExcelInteropHelper.ReleaseComObject(sourceWorkbook)
                    End Try
                End If

                If targetWorkbook IsNot Nothing Then
                    Try
                        targetWorkbook.Close(SaveChanges:=False)
                    Catch
                        ' Continue cleanup if Excel rejects the close request.
                    Finally
                        ExcelInteropHelper.ReleaseComObject(targetWorkbook)
                    End Try
                End If

                If excel IsNot Nothing Then
                    Try
                        excel.Quit()
                    Catch
                        ' Continue cleanup if Excel rejects the quit request.
                    Finally
                        ExcelInteropHelper.ReleaseComObject(excel)
                    End Try
                End If
            End Try
        End Using
    End Sub

End Class