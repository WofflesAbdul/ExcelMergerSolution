
Imports System.Threading
Imports Microsoft.Office.Interop

Public Class FunctionalTestSorter

    Public Sub SortSheets(
        filePath As String,
        Optional reportProgress As Action(Of Integer) = Nothing,
        Optional reportStatus As Action(Of String) = Nothing)

        If Thread.CurrentThread.GetApartmentState() <> ApartmentState.STA Then
            Throw New InvalidOperationException(
                "Excel sorting must run on an STA thread.")
        End If

        Dim excelApp As Excel.Application = Nothing
        Dim wb As Excel.Workbook = Nothing
        Dim messageFilter As ExcelHelper.ExcelMessageFilterScope = Nothing
        Dim completed As Boolean = False

        Try
            reportProgress?.Invoke(0)
            reportStatus?.Invoke("Opening workbook...")

            messageFilter = New ExcelHelper.ExcelMessageFilterScope()

            excelApp = New Excel.Application()
            wb = excelApp.Workbooks.Open(filePath)

            reportProgress?.Invoke(5)

            SortSheets(wb, reportProgress, reportStatus, progressStart:=5, progressEnd:=80)

            reportStatus?.Invoke("Resetting worksheet cursors...")
            ExcelHelper.ResetWorksheetCursors(wb, reportProgress, 80, 94)

            reportProgress?.Invoke(94)
            reportStatus?.Invoke("Saving workbook...")
            wb.Save()

            reportProgress?.Invoke(97)
            reportStatus?.Invoke("Closing workbook...")

            wb.Close(SaveChanges:=False)
            ExcelHelper.ReleaseComObject(wb)
            wb = Nothing

            excelApp.Quit()
            ExcelHelper.ReleaseComObject(excelApp)
            excelApp = Nothing

            completed = True

        Catch ex As Exception
            Throw New ApplicationException($"Sort failed: {ex.Message}", ex)

        Finally
            If wb IsNot Nothing Then
                Try
                    wb.Close(SaveChanges:=False)
                Catch
                End Try

                ExcelHelper.ReleaseComObject(wb)
            End If

            If excelApp IsNot Nothing Then
                Try
                    excelApp.Quit()
                Catch
                End Try

                ExcelHelper.ReleaseComObject(excelApp)
            End If

            If messageFilter IsNot Nothing Then
                messageFilter.Dispose()
            End If

            If completed Then
                reportProgress?.Invoke(100)
            End If
        End Try
    End Sub

    Public Sub SortSheets(
        wb As Excel.Workbook,
        Optional reportProgress As Action(Of Integer) = Nothing,
        Optional reportStatus As Action(Of String) = Nothing,
        Optional progressStart As Integer = 5,
        Optional progressEnd As Integer = 80)

        If wb Is Nothing Then
            Throw New ArgumentNullException(NameOf(wb))
        End If

        Dim sheetInfoList As New List(Of SheetNameParsingInfo)
        Dim sheets As Excel.Sheets = Nothing

        Try
            sheets = wb.Worksheets

            Dim totalSheets As Integer = sheets.Count
            Dim progressRange As Integer = progressEnd - progressStart
            Dim readEnd As Integer = progressStart + CInt(Math.Floor(progressRange * 0.1))
            Dim moveStart As Integer = progressStart + CInt(Math.Floor(progressRange * 0.15))

            ' Stage 1: Read worksheet names and calculate sort keys.
            reportStatus?.Invoke("Reading worksheets...")

            For i As Integer = 1 To totalSheets
                Dim ws As Excel.Worksheet = Nothing

                Try
                    ws = CType(sheets.Item(i), Excel.Worksheet)

                    Dim fullName As String = ws.Name
                    Dim allParts() As String = fullName.Split(" "c)
                    Dim parts() As String

                    If allParts.Length > 1 Then
                        parts = {
                            allParts(0),
                            String.Join(" ", allParts, 1, allParts.Length - 1)
                        }
                    Else
                        parts = allParts
                    End If

                    Dim suffix As String =
                        If(parts.Length > 1, parts(1), String.Empty)

                    Dim baseIndex As Integer =
                        FunctionalTestSequence.TestNames.FindIndex(
                            Function(b) fullName.IndexOf(
                                b,
                                StringComparison.OrdinalIgnoreCase) >= 0)

                    If baseIndex = -1 Then
                        baseIndex = Integer.MaxValue
                    End If

                    Dim voltageOutputIndex As Integer =
                        VoltageOutputSequence.GetVoltageOutputIndex(suffix)

                    Dim vpsuIndex As Integer =
                        VpsuSequence.GetVpsuIndex(suffix)

                    Dim tempIndex As Integer =
                        TemperatureSequence.GetTemperatureIndex(suffix)

                    Dim headerPriority As Integer =
                        HeaderSequence.GetHeaderPriority(fullName)

                    sheetInfoList.Add(New SheetNameParsingInfo With {
                        .Worksheet = ws,
                        .OriginalName = fullName,
                        .baseIndex = baseIndex,
                        .voltageOutputIndex = voltageOutputIndex,
                        .vpsuIndex = vpsuIndex,
                        .tempIndex = tempIndex,
                        .headerPriority = headerPriority
                    })

                    ' The list now owns the worksheet reference.
                    ws = Nothing

                Finally
                    ExcelHelper.ReleaseComObject(ws)
                End Try

                ReportStageProgress(
                    reportProgress,
                    i,
                    totalSheets,
                    progressStart,
                    readEnd)
            Next

            ' Stage 2: Sort the worksheet references in memory.
            reportStatus?.Invoke("Ordering worksheets...")
            reportProgress?.Invoke(readEnd)

            Dim sortedSheets = sheetInfoList _
                .OrderBy(Function(s) s.headerPriority) _
                .ThenBy(Function(s) s.tempIndex) _
                .ThenBy(Function(s) s.baseIndex) _
                .ThenBy(Function(s) s.voltageOutputIndex) _
                .ThenBy(Function(s) s.vpsuIndex) _
                .ToList()

            reportProgress?.Invoke(moveStart)

            ' Stage 3: Moving worksheets is expected to take most of the time.
            reportStatus?.Invoke("Moving worksheets...")

            If sortedSheets.Count > 0 Then
                Dim firstSheet As Excel.Worksheet = Nothing

                Try
                    firstSheet = CType(sheets.Item(1), Excel.Worksheet)

                    If Not String.Equals(
                        firstSheet.Name,
                        sortedSheets(0).OriginalName,
                        StringComparison.Ordinal) Then

                        sortedSheets(0).Worksheet.Move(Before:=firstSheet)
                    End If

                Finally
                    ExcelHelper.ReleaseComObject(firstSheet)
                End Try

                ReportStageProgress(
                    reportProgress,
                    1,
                    sortedSheets.Count,
                    moveStart,
                    progressEnd)

                For i As Integer = 1 To sortedSheets.Count - 1
                    sortedSheets(i).Worksheet.Move(
                        After:=sortedSheets(i - 1).Worksheet)

                    ReportStageProgress(
                        reportProgress,
                        i + 1,
                        sortedSheets.Count,
                        moveStart,
                        progressEnd)
                Next
            Else
                reportProgress?.Invoke(progressEnd)
            End If

        Catch ex As Exception
            Throw New ApplicationException($"Sort failed: {ex.Message}", ex)

        Finally
            For Each sheetInfo In sheetInfoList
                ExcelHelper.ReleaseComObject(sheetInfo.Worksheet)
            Next

            ExcelHelper.ReleaseComObject(sheets)
        End Try
    End Sub

    Private Sub ReportStageProgress(
        reportProgress As Action(Of Integer),
        processedItems As Integer,
        totalItems As Integer,
        progressStart As Integer,
        progressEnd As Integer)

        If reportProgress Is Nothing OrElse totalItems <= 0 Then
            Return
        End If

        Dim progress As Integer =
            progressStart +
            CInt(Math.Floor(
                CDbl(processedItems) / totalItems *
                (progressEnd - progressStart)))

        reportProgress.Invoke(Math.Min(progress, progressEnd))
    End Sub

End Class
