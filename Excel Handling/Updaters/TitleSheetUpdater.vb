Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop

Public Class TitleSheetUpdater

    Public Sub UpdateTitleSheet(filePath As String, values As ResolvedTestMetadata)
        Dim excelApp As Excel.Application = Nothing
        Dim wb As Excel.Workbook = Nothing

        Try
            excelApp = New Excel.Application()
            wb = excelApp.Workbooks.Open(filePath)

            Dim titleSheet As Excel.Worksheet = Nothing

            For Each ws As Excel.Worksheet In wb.Sheets
                If ws.Name.Equals("Title", StringComparison.OrdinalIgnoreCase) Then
                    titleSheet = ws
                    Exit For
                End If
            Next

            If titleSheet Is Nothing Then Return ' No Title sheet → silently exit

            WriteNamed(titleSheet, "PowerSupplyModel", values.ModelNumber)
            WriteNamed(titleSheet, "PowerSupplySerialNumber", values.SerialNumber)
            WriteNamed(titleSheet, "PowerSupplyFirmwareVersion", values.FirmwareVersion)
            'WriteNamed(titleSheet, "Description", values.DevelopmentPhase)
            'WriteNamed(titleSheet, "Tester", values.TestedBy)

            wb.Save()

        Finally
            If wb IsNot Nothing Then wb.Close(SaveChanges:=True)
            If excelApp IsNot Nothing Then excelApp.Quit()
            ReleaseComObject(wb)
            ReleaseComObject(excelApp)
        End Try
    End Sub

    Public Sub UpdateTitleSheetFromOpenWorkbook(wb As Excel.Workbook, values As ResolvedTestMetadata)
        Dim titleSheet As Excel.Worksheet = Nothing

        For Each ws As Excel.Worksheet In wb.Sheets
            If ws.Name.Equals("Title", StringComparison.OrdinalIgnoreCase) Then
                titleSheet = ws
                Exit For
            End If
        Next

        If titleSheet Is Nothing Then Return ' No Title sheet → silently exit

        WriteNamed(titleSheet, "PowerSupplyModel", values.ModelNumber)
        WriteNamed(titleSheet, "PowerSupplySerialNumber", values.SerialNumber)
        WriteNamed(titleSheet, "PowerSupplyFirmwareVersion", values.FirmwareVersion)
        'WriteNamed(titleSheet, "Description", values.DevelopmentPhase)
        'WriteNamed(titleSheet, "Tester", values.TestedBy)

        wb.Save()
    End Sub

    Private Sub WriteNamed(ws As Excel.Worksheet, name As String, value As String)
        If String.IsNullOrWhiteSpace(value) Then Return

        Try
            Dim nm = ws.Names.Item(name)
            Dim rng = DirectCast(nm.RefersToRange, Excel.Range)
            rng.Value2 = value
        Catch
            ' Named range does not exist → ignore
        End Try
    End Sub

    Private Sub WriteToTable(ws As Excel.Worksheet, tableName As String, values As IDictionary(Of String, Object), mode As TableWriteMode, Optional rowIndex As Integer = -1)
        If values Is Nothing OrElse values.Count = 0 Then Return

        Try
            Dim tbl As Excel.ListObject = ws.ListObjects(tableName)
            Dim targetRow As Excel.ListRow = Nothing

            Select Case mode
                Case TableWriteMode.AppendNew
                    targetRow = tbl.ListRows.Add()

                Case TableWriteMode.OverwriteLast
                    If tbl.ListRows.Count > 0 Then
                        targetRow = tbl.ListRows(tbl.ListRows.Count)
                    Else
                        targetRow = tbl.ListRows.Add()
                    End If

                Case TableWriteMode.OverwriteFirst
                    If tbl.ListRows.Count > 0 Then
                        targetRow = tbl.ListRows(1)
                    Else
                        targetRow = tbl.ListRows.Add()
                    End If

                Case TableWriteMode.OverwriteByIndex
                    If rowIndex > 0 AndAlso rowIndex <= tbl.ListRows.Count Then
                        targetRow = tbl.ListRows(rowIndex)
                    Else
                        targetRow = tbl.ListRows.Add()
                    End If
            End Select

            If targetRow Is Nothing Then Return

            ' Write values
            For Each kvp In values
                Dim columnName = kvp.Key
                Dim columnValue = kvp.Value

                If columnValue Is Nothing Then Continue For
                If TypeOf columnValue Is String AndAlso String.IsNullOrWhiteSpace(columnValue.ToString()) Then Continue For

                Try
                    Dim colIndex = tbl.ListColumns(columnName).Index
                    targetRow.Range(1, colIndex).Value2 = columnValue
                Catch
                    ' Column missing → ignore
                End Try
            Next

        Catch
            ' Table missing → ignore
        End Try
    End Sub


    Private Sub ReleaseComObject(obj As Object)
        Try
            If obj IsNot Nothing Then Marshal.ReleaseComObject(obj)
        Finally
            obj = Nothing
        End Try
    End Sub
End Class


