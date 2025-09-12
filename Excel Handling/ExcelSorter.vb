Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices

Public Class FunctionalTestSorter

    Public Sub SortSheets(filePath As String)
        Dim excelApp As Excel.Application = Nothing
        Dim wb As Excel.Workbook = Nothing

        Try
            excelApp = New Excel.Application()
            wb = excelApp.Workbooks.Open(filePath)

            Dim sheetInfoList As New List(Of SheetInfo)

            For Each ws As Excel.Worksheet In wb.Sheets
                Dim fullName As String = ws.Name

                ' Split into base name + suffix
                Dim allParts() As String = fullName.Split(" "c)
                Dim parts() As String
                If allParts.Length > 1 Then
                    parts = {allParts(0), String.Join(" ", allParts, 1, allParts.Length - 1)}
                Else
                    parts = allParts
                End If

                Dim baseName As String = parts(0)
                Dim suffix As String = If(parts.Length > 1, parts(1), String.Empty)

                ' Determine indices
                Dim baseIndex As Integer = FunctionalTestSequence.TestNames.FindIndex(Function(b) fullName.IndexOf(b, StringComparison.OrdinalIgnoreCase) >= 0)
                If baseIndex = -1 Then baseIndex = Integer.MaxValue

                Dim vpsuIndex As Integer = VpsuSequence.GetVpsuIndex(suffix)
                Dim tempIndex As Integer = TemperatureSequence.GetTemperatureIndex(suffix)

                sheetInfoList.Add(New SheetInfo With {
                    .Worksheet = ws,
                    .OriginalName = fullName,
                    .baseIndex = baseIndex,
                    .vpsuIndex = vpsuIndex,
                    .tempIndex = tempIndex
                })
            Next

            ' Sort by baseName → Vpsu → Temperature
            Dim sortedSheets = sheetInfoList.OrderBy(Function(s) s.baseIndex) _
                                            .ThenBy(Function(s) s.vpsuIndex) _
                                            .ThenBy(Function(s) s.tempIndex) _
                                            .ToList()

            ' Move sheets safely
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