Imports System.IO
Imports Microsoft.Office.Interop

Public NotInheritable Class ExcelWorkbookCreator

    Private Shared ReadOnly PlaceholderNames As String() = {"Placeholder", "Sheet1", "Sheet2", "Sheet3"}

    Private Sub New()
    End Sub

    Public Shared Function CreateNewExcel(directoryPath As String, fileName As String) As String
        If String.IsNullOrWhiteSpace(directoryPath) Then Throw New ArgumentException("Directory path cannot be empty.", NameOf(directoryPath))
        If String.IsNullOrWhiteSpace(fileName) Then Throw New ArgumentException("Filename cannot be empty.", NameOf(fileName))

        If Path.GetExtension(fileName).ToLowerInvariant() <> ".xlsx" Then
            fileName &= ".xlsx"
        End If

        If Not Directory.Exists(directoryPath) Then
            Directory.CreateDirectory(directoryPath)
        End If

        Dim fullPath As String = Path.Combine(directoryPath, fileName)
        Dim excelApp As Excel.Application = Nothing
        Dim workbooks As Excel.Workbooks = Nothing
        Dim wb As Excel.Workbook = Nothing

        Try
            Using filterScope As New ExcelInteropHelper.ExcelMessageFilterScope()
                excelApp = New Excel.Application()
                excelApp.DisplayAlerts = False

                workbooks = excelApp.Workbooks
                wb = workbooks.Add()

                wb.SaveAs(fullPath)
                wb.Close(SaveChanges:=False)
                ExcelInteropHelper.ReleaseComObject(wb)
                wb = Nothing
            End Using

            If Not File.Exists(fullPath) Then
                Throw New ApplicationException($"Failed to create Excel file at {fullPath}")
            End If

            Return fullPath

        Finally
            If wb IsNot Nothing Then
                Try
                    wb.Close(SaveChanges:=False)
                Catch
                    ' Preserve the original exception.
                Finally
                    ExcelInteropHelper.ReleaseComObject(wb)
                End Try
            End If

            ExcelInteropHelper.ReleaseComObject(workbooks)

            If excelApp IsNot Nothing Then
                Try
                    excelApp.Quit()
                Catch
                    ' Preserve the original exception.
                Finally
                    ExcelInteropHelper.ReleaseComObject(excelApp)
                End Try
            End If
        End Try
    End Function

    Public Shared Sub RemovePlaceholderSheets(wb As Excel.Workbook)
        If wb Is Nothing Then Throw New ArgumentNullException(NameOf(wb))

        Dim sheets As Excel.Sheets = Nothing

        Try
            sheets = wb.Worksheets

            For i As Integer = sheets.Count To 1 Step -1
                Dim ws As Excel.Worksheet = Nothing

                Try
                    ws = CType(sheets.Item(i), Excel.Worksheet)

                    If Array.Exists(PlaceholderNames, Function(name) String.Equals(name, ws.Name, StringComparison.OrdinalIgnoreCase)) Then
                        ws.Delete()
                    End If

                Finally
                    ExcelInteropHelper.ReleaseComObject(ws)
                End Try
            Next

        Finally
            ExcelInteropHelper.ReleaseComObject(sheets)
        End Try
    End Sub

End Class