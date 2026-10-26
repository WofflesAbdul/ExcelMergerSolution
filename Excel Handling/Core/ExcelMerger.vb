Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop

Public Class ExcelMerger

    Public Sub MergeFiles(baseFile As String, targets As IEnumerable(Of String), Optional reportProgress As Action(Of Integer) = Nothing)
        Dim excelApp As Excel.Application = Nothing
        Dim destWB As Excel.Workbook = Nothing

        Try
            excelApp = New Excel.Application()
            excelApp.DisplayAlerts = False
            destWB = excelApp.Workbooks.Open(baseFile)

            MergeSheets(excelApp, destWB, targets, reportProgress)
            ExcelFileCreator.RemovePlaceholderSheets(destWB)

            destWB.Save()

        Catch ex As Exception
            Throw New ApplicationException($"Merge failed: {ex.Message}", ex)

        Finally
            If destWB IsNot Nothing Then
                destWB.Close(SaveChanges:=False)
                ReleaseComObject(destWB)
            End If

            If excelApp IsNot Nothing Then
                excelApp.Quit()
                ReleaseComObject(excelApp)
            End If
        End Try
    End Sub

    Public Sub MergeAndSortFiles(baseFile As String, targets As IEnumerable(Of String), Optional reportProgress As Action(Of Integer) = Nothing)
        Dim excelApp As Excel.Application = Nothing
        Dim destWB As Excel.Workbook = Nothing

        Try
            excelApp = New Excel.Application()
            excelApp.DisplayAlerts = False
            destWB = excelApp.Workbooks.Open(baseFile)

            ' Merge worksheets into the existing destination workbook.
            MergeSheets(excelApp, destWB, targets, reportProgress)

            ' Remove placeholders without reopening the destination workbook.
            ExcelFileCreator.RemovePlaceholderSheets(destWB)

            ' Sort worksheets in the same workbook session.
            Dim sorter As New FunctionalTestSorter()
            sorter.SortSheets(destWB)

            destWB.Save()

        Catch ex As Exception
            Throw New ApplicationException($"Merge and sort failed: {ex.Message}", ex)

        Finally
            If destWB IsNot Nothing Then
                destWB.Close(SaveChanges:=False)
                ReleaseComObject(destWB)
            End If

            If excelApp IsNot Nothing Then
                excelApp.Quit()
                ReleaseComObject(excelApp)
            End If
        End Try
    End Sub

    Public Sub MergeSheets(excelApp As Excel.Application, destWB As Excel.Workbook, targets As IEnumerable(Of String), Optional reportProgress As Action(Of Integer) = Nothing)
        If excelApp Is Nothing Then Throw New ArgumentNullException(NameOf(excelApp))
        If destWB Is Nothing Then Throw New ArgumentNullException(NameOf(destWB))
        If targets Is Nothing Then Throw New ArgumentNullException(NameOf(targets))

        Dim targetFiles As List(Of String) = targets.ToList()
        Dim total As Integer = targetFiles.Count
        Dim index As Integer = 0

        For Each filePath As String In targetFiles
            Dim sourceWB As Excel.Workbook = Nothing
            Dim wsToCopy As Excel.Worksheet = Nothing
            Dim copiedSheet As Excel.Worksheet = Nothing
            Dim sheets As Excel.Sheets = Nothing

            Try
                sourceWB = excelApp.Workbooks.Open(filePath)
                wsToCopy = CType(sourceWB.Sheets(1), Excel.Worksheet)

                Dim newSheetName As String = ResolveSheetNameConflict(wsToCopy.Name, destWB)

                sheets = destWB.Sheets
                wsToCopy.Copy(After:=sheets(sheets.Count))

                copiedSheet = CType(sheets(sheets.Count), Excel.Worksheet)
                copiedSheet.Name = newSheetName

            Finally
                If sourceWB IsNot Nothing Then
                    sourceWB.Close(SaveChanges:=False)
                End If

                ReleaseComObject(copiedSheet)
                ReleaseComObject(wsToCopy)
                ReleaseComObject(sheets)
                ReleaseComObject(sourceWB)
            End Try

            index += 1

            If total > 0 Then
                reportProgress?.Invoke(CInt(index * 100 / total))
            End If
        Next
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
        Dim sheets As Excel.Sheets = Nothing

        Try
            sheets = wb.Sheets

            For i As Integer = 1 To sheets.Count
                Dim ws As Excel.Worksheet = Nothing

                Try
                    ws = CType(sheets(i), Excel.Worksheet)

                    If String.Equals(ws.Name, sheetName, StringComparison.OrdinalIgnoreCase) Then
                        Return True
                    End If
                Finally
                    ReleaseComObject(ws)
                End Try
            Next

            Return False
        Finally
            ReleaseComObject(sheets)
        End Try
    End Function

    Private Sub ReleaseComObject(obj As Object)
        Try
            If obj IsNot Nothing AndAlso Marshal.IsComObject(obj) Then
                Marshal.ReleaseComObject(obj)
            End If
        Catch
        End Try
    End Sub
End Class
