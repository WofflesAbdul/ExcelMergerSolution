Imports System.IO
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop

Public NotInheritable Class ExcelFileCreator

    Private Shared ReadOnly PlaceholderNames As String() = {"Placeholder", "Sheet1", "Sheet2", "Sheet3"}

    Public Shared Function CreateNewExcel(directoryPath As String, fileName As String) As String
        If String.IsNullOrWhiteSpace(directoryPath) Then Throw New ArgumentException("Directory path cannot be empty.", NameOf(directoryPath))
        If String.IsNullOrWhiteSpace(fileName) Then Throw New ArgumentException("Filename cannot be empty.", NameOf(fileName))

        If Path.GetExtension(fileName).ToLowerInvariant() <> ".xlsx" Then
            fileName &= ".xlsx"
        End If

        Dim fullPath As String = Path.Combine(directoryPath, fileName)

        If Not Directory.Exists(directoryPath) Then
            Directory.CreateDirectory(directoryPath)
        End If

        Dim excelApp As Excel.Application = Nothing
        Dim wb As Excel.Workbook = Nothing

        Try
            excelApp = New Excel.Application()
            excelApp.DisplayAlerts = False

            wb = excelApp.Workbooks.Add()
            wb.SaveAs(fullPath)
            wb.Close(SaveChanges:=False)
            ReleaseComObject(wb)
            wb = Nothing

            If Not File.Exists(fullPath) Then
                Throw New ApplicationException($"Failed to create Excel file at {fullPath}")
            End If

            Return fullPath

        Finally
            If wb IsNot Nothing Then
                wb.Close(SaveChanges:=False)
                ReleaseComObject(wb)
            End If

            If excelApp IsNot Nothing Then
                excelApp.Quit()
                ReleaseComObject(excelApp)
            End If
        End Try
    End Function

    Public Shared Sub RemovePlaceholderSheets(fullPath As String)
        If Not File.Exists(fullPath) Then Return

        Dim excelApp As Excel.Application = Nothing
        Dim wb As Excel.Workbook = Nothing

        Try
            excelApp = New Excel.Application()
            excelApp.DisplayAlerts = False
            wb = excelApp.Workbooks.Open(fullPath)

            RemovePlaceholderSheets(wb)

            wb.Save()

        Catch ex As Exception
            Throw New ApplicationException($"Failed to remove placeholder sheets: {ex.Message}", ex)

        Finally
            If wb IsNot Nothing Then
                wb.Close(SaveChanges:=False)
                ReleaseComObject(wb)
            End If

            If excelApp IsNot Nothing Then
                excelApp.Quit()
                ReleaseComObject(excelApp)
            End If
        End Try
    End Sub

    Public Shared Sub RemovePlaceholderSheets(wb As Excel.Workbook)
        If wb Is Nothing Then Throw New ArgumentNullException(NameOf(wb))

        Dim sheets As Excel.Sheets = Nothing

        Try
            sheets = wb.Sheets

            For i As Integer = sheets.Count To 1 Step -1
                Dim ws As Excel.Worksheet = Nothing

                Try
                    ws = CType(sheets(i), Excel.Worksheet)

                    If PlaceholderNames.Contains(ws.Name, StringComparer.OrdinalIgnoreCase) Then
                        ws.Delete()
                    End If

                Finally
                    ReleaseComObject(ws)
                End Try
            Next

        Finally
            ReleaseComObject(sheets)
        End Try
    End Sub

    Private Shared Sub ReleaseComObject(obj As Object)
        Try
            If obj IsNot Nothing AndAlso Marshal.IsComObject(obj) Then
                Marshal.ReleaseComObject(obj)
            End If
        Catch
        End Try
    End Sub
End Class
