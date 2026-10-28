Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Threading
Imports Excel = Microsoft.Office.Interop.Excel

''' <summary>
''' Handles report template operations on Excel workbooks.
''' </summary>
Public Class ExcelWorksheetAdderCoverPage

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
        Dim headerInserter As New ExcelWorksheetHeaderInserter()

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
            headerInserter.CopyHeaderWorksheet(coverSheet, targetSheets, coverSheetName)

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

End Class
