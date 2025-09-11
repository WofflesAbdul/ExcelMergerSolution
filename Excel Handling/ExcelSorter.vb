Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices

Public Class ExcelSorter

    Public Sub SortSheets(filePath As String)
        Dim excelApp As Excel.Application = Nothing
        Dim wb As Excel.Workbook = Nothing

        Try
            excelApp = New Excel.Application()
            wb = excelApp.Workbooks.Open(filePath)

            Dim sheetInfoList As New List(Of SheetInfo)

            For Each ws As Excel.Worksheet In wb.Sheets
                Dim fullName As String = ws.Name
                Dim baseIndex As Integer = Integer.MaxValue

                ' Determine base index using Sequence1 order
                For i As Integer = 0 To Sequence1.BaseNames.Count - 1
                    If fullName.IndexOf(Sequence1.BaseNames(i), StringComparison.OrdinalIgnoreCase) >= 0 Then
                        baseIndex = i
                        Exit For
                    End If
                Next

                sheetInfoList.Add(New SheetInfo With {
                    .Worksheet = ws,
                    .OriginalName = fullName,
                    .BaseIndex = baseIndex
                })
            Next

            ' Sort sheets: primary = Sequence1 order, secondary = natural/alphabetical
            Dim sortedSheets = sheetInfoList _
                               .OrderBy(Function(s) s.BaseIndex) _
                               .ThenBy(Function(s) s.OriginalName, StringComparer.OrdinalIgnoreCase) _
                               .ToList()

            ' Move sheets in reverse to preserve final order
            For i As Integer = sortedSheets.Count - 1 To 0 Step -1
                sortedSheets(i).Worksheet.Move(Before:=wb.Sheets(1))
            Next

            wb.Save()

        Catch ex As Exception
            Throw New ApplicationException($"Sort failed: {ex.Message}", ex)
        Finally
            If wb IsNot Nothing Then wb.Close(SaveChanges:=True)
            If excelApp IsNot Nothing Then excelApp.Quit()

            ReleaseComObject(wb)
            ReleaseComObject(excelApp)
        End Try
    End Sub

    Private Sub ReleaseComObject(obj As Object)
        Try
            If obj IsNot Nothing Then Marshal.ReleaseComObject(obj)
        Catch
        Finally
            obj = Nothing
        End Try
    End Sub

End Class
