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
            WriteNamed(titleSheet, "Description", values.DevelopmentPhase)
            WriteNamed(titleSheet, "Tester", values.TestedBy)

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
        WriteNamed(titleSheet, "Description", values.DevelopmentPhase)
        WriteNamed(titleSheet, "Tester", values.TestedBy)

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

    Private Sub ReleaseComObject(obj As Object)
        Try
            If obj IsNot Nothing Then Marshal.ReleaseComObject(obj)
        Finally
            obj = Nothing
        End Try
    End Sub
End Class


