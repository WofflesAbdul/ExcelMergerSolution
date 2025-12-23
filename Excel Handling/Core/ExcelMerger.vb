Imports Microsoft.Office.Interop

Public Class ExcelMerger
    Public Sub MergeFiles(baseFile As String, targets As IEnumerable(Of String), Optional reportProgress As Action(Of Integer) = Nothing)
        Dim excelApp As Excel.Application = Nothing
        Dim destWB As Excel.Workbook = Nothing

        Try
            excelApp = New Excel.Application()
            destWB = excelApp.Workbooks.Open(baseFile)

            Dim total As Integer = targets.Count()
            Dim index As Integer = 0

            For Each filePath In targets
                Dim sourceWB As Excel.Workbook = excelApp.Workbooks.Open(filePath)
                Dim wsToCopy As Excel.Worksheet = sourceWB.Sheets(1)

                ' Resolve conflicts
                Dim newSheetName As String = ResolveSheetNameConflict(wsToCopy.Name, destWB)

                wsToCopy.Copy(After:=destWB.Sheets(destWB.Sheets.Count))
                destWB.Sheets(destWB.Sheets.Count).Name = newSheetName

                sourceWB.Close(SaveChanges:=False)

                ' Report progress
                index += 1
                Dim percent As Integer = CInt(index * 100 / total)
                reportProgress?.Invoke(percent)
            Next

            destWB.Save()

            Dim collector As New WorkbookNamedRangeCollector()
            Dim data = collector.CollectFromOpenWorkbook(destWB)
            Dim resolved As ResolvedTestMetadata = TestMetadataProcessor.ResolveDominant(data.TestSheets)
            Dim updater As New TitleSheetUpdater()
            updater.UpdateTitleSheetFromOpenWorkbook(destWB, resolved)

        Catch ex As Exception
            Throw New ApplicationException($"Merge failed: {ex.Message}", ex)

        Finally
            If destWB IsNot Nothing Then destWB.Close(SaveChanges:=True)
            If excelApp IsNot Nothing Then excelApp.Quit()
            ReleaseComObject(destWB)
            ReleaseComObject(excelApp)
        End Try
    End Sub

    Private Function ResolveSheetNameConflict(sheetName As String, wb As Excel.Workbook) As String
        Dim newName As String = sheetName
        Dim counter As Integer = 1
        Dim maxLength As Integer = 31
        Dim truncatedName As String = If(sheetName.Length > maxLength - 4, sheetName.Substring(0, maxLength - 4), sheetName)

        While SheetExists(newName, wb)
            newName = truncatedName & "-" & counter.ToString("00")
            counter += 1
        End While

        Return newName
    End Function

    Private Function SheetExists(sheetName As String, wb As Excel.Workbook) As Boolean
        For Each ws As Excel.Worksheet In wb.Sheets
            If ws.Name = sheetName Then Return True
        Next
        Return False
    End Function

    Private Sub ReleaseComObject(ByVal obj As Object)
        Try
            If obj IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            End If
        Finally
            obj = Nothing
        End Try
    End Sub
End Class
