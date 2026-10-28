Imports System.Threading
Imports Microsoft.Office.Interop

Public Class ExcelWorkbookMerger

    Public Sub MergeFiles(baseFile As String, targets As IEnumerable(Of String), Optional reportProgress As Action(Of Integer) = Nothing, Optional reportStatus As Action(Of String) = Nothing)
        If Thread.CurrentThread.GetApartmentState() <> ApartmentState.STA Then
            Throw New InvalidOperationException("Excel merging must run on an STA thread.")
        End If

        Dim excelApp As Excel.Application = Nothing
        Dim destWB As Excel.Workbook = Nothing
        Dim messageFilter As ExcelInteropHelper.ExcelMessageFilterScope = Nothing

        Try
            messageFilter = New ExcelInteropHelper.ExcelMessageFilterScope()

            reportStatus?.Invoke("Opening workbook...")
            reportProgress?.Invoke(0)

            excelApp = New Excel.Application()
            excelApp.DisplayAlerts = False
            destWB = excelApp.Workbooks.Open(baseFile)

            reportProgress?.Invoke(5)
            reportStatus?.Invoke("Merging files...")

            MergeSheets(excelApp, destWB, targets, reportProgress, 5, 85)
            ExcelWorkbookCreator.RemovePlaceholderSheets(destWB)

            reportStatus?.Invoke("Resetting worksheet cursors...")
            ExcelInteropHelper.ResetWorksheetCursors(destWB, reportProgress, 85, 94)

            reportProgress?.Invoke(94)
            reportStatus?.Invoke("Saving workbook...")
            destWB.Save()

            reportStatus?.Invoke("Closing workbook...")
            reportProgress?.Invoke(97)
            destWB.Close(SaveChanges:=False)
            ExcelInteropHelper.ReleaseComObject(destWB)
            destWB = Nothing

            excelApp.Quit()
            ExcelInteropHelper.ReleaseComObject(excelApp)
            excelApp = Nothing

            reportProgress?.Invoke(100)

        Catch ex As Exception
            Throw New ApplicationException($"Merge failed: {ex.Message}", ex)

        Finally
            If destWB IsNot Nothing Then
                Try
                    destWB.Close(SaveChanges:=False)
                Catch
                End Try

                ExcelInteropHelper.ReleaseComObject(destWB)
            End If

            If excelApp IsNot Nothing Then
                Try
                    excelApp.Quit()
                Catch
                End Try

                ExcelInteropHelper.ReleaseComObject(excelApp)
            End If

            If messageFilter IsNot Nothing Then
                messageFilter.Dispose()
            End If
        End Try
    End Sub

    Public Sub MergeAndSortFiles(baseFile As String, targets As IEnumerable(Of String), Optional reportProgress As Action(Of Integer) = Nothing, Optional reportStatus As Action(Of String) = Nothing)
        If Thread.CurrentThread.GetApartmentState() <> ApartmentState.STA Then
            Throw New InvalidOperationException("Excel merging must run on an STA thread.")
        End If

        Dim excelApp As Excel.Application = Nothing
        Dim destWB As Excel.Workbook = Nothing
        Dim messageFilter As ExcelInteropHelper.ExcelMessageFilterScope = Nothing

        Try
            messageFilter = New ExcelInteropHelper.ExcelMessageFilterScope()

            reportStatus?.Invoke("Opening workbook...")
            reportProgress?.Invoke(0)

            excelApp = New Excel.Application()
            excelApp.DisplayAlerts = False
            destWB = excelApp.Workbooks.Open(baseFile)

            reportProgress?.Invoke(5)
            reportStatus?.Invoke("Merging files...")

            MergeSheets(excelApp, destWB, targets, reportProgress, 5, 40)
            ExcelWorkbookCreator.RemovePlaceholderSheets(destWB)

            reportStatus?.Invoke("Sorting worksheets...")

            Dim sorter As New ExcelWorkbookSorter()
            sorter.SortSheets(destWB, reportProgress, reportStatus, progressStart:=40, progressEnd:=80)

            reportStatus?.Invoke("Resetting worksheet cursors...")
            ExcelInteropHelper.ResetWorksheetCursors(destWB, reportProgress, 80, 94)

            reportProgress?.Invoke(94)
            reportStatus?.Invoke("Saving workbook...")
            destWB.Save()

            reportStatus?.Invoke("Closing workbook...")
            reportProgress?.Invoke(97)
            destWB.Close(SaveChanges:=False)
            ExcelInteropHelper.ReleaseComObject(destWB)
            destWB = Nothing

            excelApp.Quit()
            ExcelInteropHelper.ReleaseComObject(excelApp)
            excelApp = Nothing

            reportProgress?.Invoke(100)

        Catch ex As Exception
            Throw New ApplicationException($"Merge and sort failed: {ex.Message}", ex)

        Finally
            If destWB IsNot Nothing Then
                Try
                    destWB.Close(SaveChanges:=False)
                Catch
                End Try

                ExcelInteropHelper.ReleaseComObject(destWB)
            End If

            If excelApp IsNot Nothing Then
                Try
                    excelApp.Quit()
                Catch
                End Try

                ExcelInteropHelper.ReleaseComObject(excelApp)
            End If

            If messageFilter IsNot Nothing Then
                messageFilter.Dispose()
            End If
        End Try
    End Sub

    Public Sub MergeSheets(excelApp As Excel.Application, destWB As Excel.Workbook, targets As IEnumerable(Of String), Optional reportProgress As Action(Of Integer) = Nothing, Optional progressStart As Integer = 0, Optional progressEnd As Integer = 100)
        If excelApp Is Nothing Then Throw New ArgumentNullException(NameOf(excelApp))
        If destWB Is Nothing Then Throw New ArgumentNullException(NameOf(destWB))
        If targets Is Nothing Then Throw New ArgumentNullException(NameOf(targets))

        Dim targetFiles As List(Of String) = targets.ToList()
        Dim total As Integer = targetFiles.Count
        Dim index As Integer = 0

        If total = 0 Then
            reportProgress?.Invoke(progressEnd)
            Return
        End If

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

            Catch ex As Exception
                Throw New ApplicationException($"Failed to merge target file '{filePath}': {ex.Message}", ex)

            Finally
                If sourceWB IsNot Nothing Then
                    Try
                        sourceWB.Close(SaveChanges:=False)
                    Catch
                    End Try
                End If

                ExcelInteropHelper.ReleaseComObject(copiedSheet)
                ExcelInteropHelper.ReleaseComObject(wsToCopy)
                ExcelInteropHelper.ReleaseComObject(sheets)
                ExcelInteropHelper.ReleaseComObject(sourceWB)
            End Try

            index += 1

            Dim percent As Integer = progressStart + CInt(Math.Floor(index * (progressEnd - progressStart) / CDbl(total)))
            reportProgress?.Invoke(percent)
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
                    ExcelInteropHelper.ReleaseComObject(ws)
                End Try
            Next

            Return False
        Finally
            ExcelInteropHelper.ReleaseComObject(sheets)
        End Try
    End Function

End Class